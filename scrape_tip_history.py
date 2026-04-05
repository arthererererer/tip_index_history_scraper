"""
TIP 臺灣指數公司 — 指數歷史 CSV 下載與合併

1. 以自訂 User-Agent 取得「各項指數」完整列表（count=-1）。
2. 使用 Playwright 開啟各指數「歷史指數值」頁、填日期、搜尋、下載。
3. 可單一指數輸出，或以 --all 將全部指數合併為同一 CSV（列前加上指數代碼、指數名稱）。

需先執行：python -m playwright install chromium
"""

from __future__ import annotations

import argparse
import csv
import os
from datetime import date, timedelta
import random
import re
import tempfile
from pathlib import Path
from typing import Iterable

import requests
from bs4 import BeautifulSoup
from playwright.sync_api import Page, sync_playwright

BASE = "https://taiwanindex.com.tw"
# 一次載入全部指數列（與官網「全部」筆數相同）
INDEX_LIST_URL = f"{BASE}/indexes?count=-1&page=1"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36"
)
DEFAULT_INDEX_NAME = "特選臺灣動能優息指數"
# 列表 href 為 /indexes/<單段> 但非指數詳情頁的 slug（導覽用）
EXCLUDED_INDEX_SLUGS = frozenset(
    {
        "comparison",
        "board",
        "categories",
        "multiple",
        "multipleHistory",
    }
)
# 避免 Windows 終端機 / 編輯器非 UTF-8 時字串毀損
BTN_SEARCH = "\u641c\u5c0b"  # 搜尋
BTN_DOWNLOAD = "\u4e0b\u8f09"  # 下載

MERGED_COLUMNS = (
    "\u6307\u6578\u4ee3\u78bc",  # 指數代碼
    "\u6307\u6578\u540d\u7a31",  # 指數名稱
    "\u65e5\u671f",  # 日期
    "\u50f9\u683c\u6307\u6578\u503c",  # 價格指數值
    "\u5831\u916c\u6307\u6578\u503c",  # 報酬指數值
    "\u6f32\u8dcc\u9ede\u6578",  # 漲跌點數
    "\u6f32\u8dcc\u767e\u5206\u6bd4",  # 漲跌百分比
)


def session_with_ua() -> requests.Session:
    s = requests.Session()
    s.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
        }
    )
    return s


def _register_index(seen: dict[str, str], code: str, name: str) -> None:
    if code in EXCLUDED_INDEX_SLUGS:
        return
    if code not in seen:
        seen[code] = (name or "").strip() or code


def list_all_indexes() -> list[tuple[str, str]]:
    """
    從 INDEX_LIST_URL 解析所有指數：(代碼, 名稱)。
    官網「全部」列表可能為桌面版 <table>，或行動版 / 長列表的 div.index-box，兩者皆解析；依代碼去重（保留第一次）。
    """
    sess = session_with_ua()
    r = sess.get(INDEX_LIST_URL, timeout=120)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    seen: dict[str, str] = {}
    href_re = re.compile(r"^/indexes/([A-Za-z0-9]+)/?$")

    # 行動版或 count=-1 長列表：div.index-box + .item-title + 內部「查看詳情」連結
    for box in soup.select("div.index-box"):
        title_el = box.select_one(".item-title")
        name = title_el.get_text(strip=True) if title_el else ""
        for a in box.find_all("a", href=True):
            m = href_re.match(a["href"])
            if not m:
                continue
            _register_index(seen, m.group(1), name)
            break

    # 桌面版表格：<tr> + td.index-name-cell
    for tr in soup.find_all("tr"):
        cells = tr.find_all("td")
        if len(cells) < 2:
            continue
        code: str | None = None
        for a in tr.find_all("a", href=True):
            m = href_re.match(a["href"])
            if not m:
                continue
            slug = m.group(1)
            if slug in EXCLUDED_INDEX_SLUGS:
                continue
            code = slug
            break
        if not code:
            continue
        name_cell = tr.select_one("td.index-name-cell")
        if name_cell:
            name = name_cell.get_text(strip=True)
        else:
            name = cells[1].get_text(strip=True)
        _register_index(seen, code, name)

    return sorted(seen.items(), key=lambda x: x[0])


def resolve_index_code(index_name: str) -> str:
    """從完整列表頁找出與指數名稱相符列的代碼（子字串比對）。"""
    for code, name in list_all_indexes():
        if index_name in name:
            return code
    raise ValueError(f"列表頁找不到指數名稱：{index_name!r}")


def _tip_date_to_iso(s: str) -> str:
    """CLI 常用 2023/01/01；官網 SSR / flatpickr 多為 2023-01-01。"""
    s = (s or "").strip().replace(".", "/")
    if "/" in s:
        parts = s.split("/")
        if len(parts) == 3:
            y, m, d = parts
            return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        return s[:10]
    return s


def _normalize_tip_date_display(s: str) -> str:
    """比對 input 顯示值用（統一成 YYYY-MM-DD）。"""
    return _tip_date_to_iso(s.replace(".", "/"))


def _apply_dates_to_history_inputs(page: Page, start_date: str, end_date: str) -> None:
    """
    歷史頁日期由 Vue / flatpickr 綁定。官網 __NUXT__ 為 YYYY-MM-DD；傳入 slash 時 flatpickr
    可能無法解析，導致 setDate 無效、搜尋仍用 SSR 預設區間。此處一律轉 ISO，並等 flatpickr
    掛載後再 setDate；必要時用鍵盤輸入備援。
    """
    start_iso = _tip_date_to_iso(start_date)
    end_iso = _tip_date_to_iso(end_date)

    # 等兩個日期欄都掛上 flatpickr（避免 hydration 前就 setDate 無效）
    try:
        page.wait_for_function(
            """() => {
                const nodes = [...document.querySelectorAll('input.rounded-input[name="date"]')];
                return nodes.length >= 2 && nodes.slice(0, 2).every((n) => n._flatpickr);
            }""",
            timeout=20_000,
        )
    except Exception:
        pass

    ok = page.evaluate(
        """([start, end]) => {
            const nodes = [...document.querySelectorAll('input.rounded-input[name="date"]')];
            if (nodes.length < 2) return false;
            const setNative = (el, v) => {
                const desc = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value'
                );
                if (desc && desc.set) desc.set.call(el, v);
                else el.value = v;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            };
            const apply = (el, iso) => {
                const fp = el._flatpickr;
                const slash = iso.replace(/-/g, '/');
                if (fp && typeof fp.setDate === 'function') {
                    let d = null;
                    try {
                        d = fp.parseDate(iso);
                    } catch (e) {}
                    if (!d) {
                        try {
                            d = fp.parseDate(slash);
                        } catch (e2) {}
                    }
                    if (d) fp.setDate(d, true);
                    else fp.setDate(iso, true);
                    if (typeof fp.redraw === 'function') fp.redraw();
                    if (fp.altInput) {
                        fp.altInput.dispatchEvent(new Event('input', { bubbles: true }));
                        fp.altInput.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    return;
                }
                setNative(el, slash);
            };
            apply(nodes[0], start);
            apply(nodes[1], end);
            return true;
        }""",
        [start_iso, end_iso],
    )
    if not ok:
        raise RuntimeError("找不到兩個日期輸入欄，無法套用起訖日期")
    page.wait_for_timeout(400)

    loc0 = page.locator('input.rounded-input[name="date"]').nth(0)
    loc1 = page.locator('input.rounded-input[name="date"]').nth(1)
    try:
        v0 = _normalize_tip_date_display(loc0.input_value(timeout=3_000))
        v1 = _normalize_tip_date_display(loc1.input_value(timeout=3_000))
    except Exception:
        v0, v1 = "", ""

    if v0 != start_iso or v1 != end_iso:
        for loc, iso in ((loc0, start_iso), (loc1, end_iso)):
            loc.click(force=True)
            loc.press("Control+A")
            loc.press("Backspace")
            loc.press_sequentially(iso, delay=35)
        page.keyboard.press("Tab")
        page.wait_for_timeout(400)


def _wait_after_history_search(page: Page) -> None:
    """
    按下「搜尋」後等表格／API 完成再下載。僅固定 sleep 容易在下載到「尚未套用區間」的
    空檔或舊資料；另補 networkidle 與簡單的 loading 消失等待。
    """
    try:
        page.wait_for_load_state("networkidle", timeout=90_000)
    except Exception:
        pass
    page.wait_for_timeout(2500)
    try:
        page.wait_for_function(
            """() => {
                const m = document.querySelector('main');
                if (!m) return true;
                const spin = m.querySelector(
                    '[class*="spinner"], [class*="loading"], [aria-busy="true"]'
                );
                if (!spin) return true;
                const st = window.getComputedStyle(spin);
                return st.display === 'none' || st.visibility === 'hidden' || spin.offsetParent === null;
            }""",
            timeout=45_000,
        )
    except Exception:
        pass
    page.wait_for_timeout(400)


def _filter_rows_by_query_range(
    rows: list[dict[str, str]], start_date: str, end_date: str
) -> list[dict[str, str]]:
    """只保留官網 CSV 中、日期落在 CLI 起訖（含）內的列；避免誤下載到整段歷史。"""
    lo = _tip_date_to_iso(start_date)
    hi = _tip_date_to_iso(end_date)
    if lo > hi:
        lo, hi = hi, lo
    out: list[dict[str, str]] = []
    for r in rows:
        raw = (r.get("日期") or "").strip()
        if not raw:
            continue
        d = _tip_date_to_iso(raw.replace(".", "/"))
        if len(d) >= 10 and lo <= d[:10] <= hi:
            out.append(r)
    return out


def _dismiss_blocking_overlays(page: Page) -> None:
    """
    搜尋後官網常以 alert-drop-shadow / alert-container 浮層提示（含查無資料），
    會擋住「下載」按鈕的點擊。先 ESC 關閉，再對仍存在的遮罩關閉 pointer-events。
    """
    for _ in range(3):
        page.keyboard.press("Escape")
        page.wait_for_timeout(150)
    page.evaluate(
        """() => {
            const sel = '.alert-drop-shadow, .alert-container, [class*="alert-drop-shadow"]';
            document.querySelectorAll(sel).forEach((el) => {
                if (!(el instanceof HTMLElement)) return;
                const r = el.getBoundingClientRect();
                if (r.width < 10 || r.height < 10) return;
                const close = el.querySelector(
                    'button[class*="close"], .btn-close, [aria-label="Close"], '
                    + '[aria-label="關閉"]'
                );
                if (close instanceof HTMLElement) close.click();
                el.style.pointerEvents = 'none';
            });
        }"""
    )
    page.wait_for_timeout(200)


def _download_history_on_page(
    page: Page,
    index_code: str,
    start_date: str,
    end_date: str,
    dest: Path,
) -> None:
    history_url = f"{BASE}/indexes/{index_code}/history"
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)

    page.goto(history_url, wait_until="networkidle", timeout=60_000)

    dates = page.locator('input.rounded-input[name="date"]')
    dates.first.wait_for(state="attached", timeout=15_000)
    if dates.count() < 2:
        raise RuntimeError(
            "歷史頁沒有兩個日期輸入欄（此代碼可能為產業類股等專用頁，與一般指數下載介面不同）"
        )
    _apply_dates_to_history_inputs(page, start_date, end_date)

    main = page.locator("main")
    main.get_by_role("button", name=BTN_SEARCH).last.click()
    _wait_after_history_search(page)
    _dismiss_blocking_overlays(page)

    with page.expect_download(timeout=60_000) as dl_info:
        main.get_by_role("button", name=BTN_DOWNLOAD).click(force=True)
    download = dl_info.value
    download.save_as(str(dest))


def scrape_history_csv(
    index_code: str,
    start_date: str,
    end_date: str,
    out_path: Path,
    headless: bool = True,
) -> Path:
    """單一指數：啟動瀏覽器、下載一個 CSV。"""
    out_path = Path(out_path)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            user_agent=USER_AGENT,
            locale="zh-TW",
            timezone_id="Asia/Taipei",
            accept_downloads=True,
        )
        page = context.new_page()
        _download_history_on_page(page, index_code, start_date, end_date, out_path)
        context.close()
        browser.close()
    return out_path


def _read_official_history_rows(csv_path: Path) -> list[dict[str, str]]:
    """讀取官網下載的 CSV（可能含 BOM），回傳資料列（不含表頭）。"""
    text = csv_path.read_text(encoding="utf-8-sig", errors="replace")
    lines = text.splitlines()
    if not lines:
        return []
    reader = csv.DictReader(lines)
    if not reader.fieldnames:
        return []
    rows: list[dict[str, str]] = []
    for row in reader:
        norm = {(k or "").strip(): (v or "").strip() for k, v in row.items()}
        if any(norm.values()):
            rows.append(norm)
    return rows


def scrape_all_merged_csv(
    start_date: str,
    end_date: str,
    out_path: Path,
    *,
    headless: bool = True,
    limit: int | None = None,
    indexes: Iterable[tuple[str, str]] | None = None,
) -> Path:
    """
    逐一抓取指數歷史並合併為單一 CSV。
    欄位：指數代碼、指數名稱、日期、價格指數值、報酬指數值、漲跌點數、漲跌百分比。
    每兩支指數之間隨機等待 1～5 秒（含），降低連線頻率。
    """
    pairs = list(indexes) if indexes is not None else list_all_indexes()
    if limit is not None:
        pairs = pairs[:limit]
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    errors: list[str] = []
    n_ok = 0
    n_rows = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            user_agent=USER_AGENT,
            locale="zh-TW",
            timezone_id="Asia/Taipei",
            accept_downloads=True,
        )
        page = context.new_page()

        with out_path.open("w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.DictWriter(out_f, fieldnames=list(MERGED_COLUMNS))
            writer.writeheader()
            out_f.flush()

            total = len(pairs)
            for i, (code, name) in enumerate(pairs):
                print(f"[{i + 1}/{total}] {code} {name}", flush=True)
                fd, tmp_name = tempfile.mkstemp(suffix=".csv")
                os.close(fd)
                tmp = Path(tmp_name)
                try:
                    _download_history_on_page(page, code, start_date, end_date, tmp)
                    rows = _read_official_history_rows(tmp)
                    rows = _filter_rows_by_query_range(rows, start_date, end_date)
                    for row in rows:
                        writer.writerow(
                            {
                                MERGED_COLUMNS[0]: code,
                                MERGED_COLUMNS[1]: name,
                                MERGED_COLUMNS[2]: (row.get("日期") or "").strip(),
                                MERGED_COLUMNS[3]: (row.get("價格指數值") or "").strip(),
                                MERGED_COLUMNS[4]: (row.get("報酬指數值") or "").strip(),
                                MERGED_COLUMNS[5]: (row.get("漲跌點數") or "").strip(),
                                MERGED_COLUMNS[6]: (row.get("漲跌百分比") or "").strip(),
                            }
                        )
                    out_f.flush()
                    n_ok += 1
                    n_rows += len(rows)
                    print(f"  -> 成功，+{len(rows)} 列", flush=True)
                except Exception as exc:  # noqa: BLE001
                    err = f"{code} ({name}): {exc}"
                    errors.append(err)
                    print(f"  -> 失敗：{exc}", flush=True)
                finally:
                    if tmp.exists():
                        tmp.unlink(missing_ok=True)

                if i + 1 < total:
                    wait_ms = int(random.uniform(1.0, 5.0) * 1000)
                    page.wait_for_timeout(wait_ms)

        context.close()
        browser.close()

    if errors:
        err_path = out_path.with_suffix(".errors.txt")
        err_path.write_text("\n".join(errors), encoding="utf-8")
        print(f"有 {len(errors)} 筆指數失敗，詳情：{err_path.resolve()}")

    print(
        f"完成：成功 {n_ok}/{len(pairs)} 支指數，合計 {n_rows} 列資料 -> {out_path.resolve()}",
        flush=True,
    )
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(description="TIP 指數歷史 CSV 下載（Playwright）")
    parser.add_argument(
        "--all",
        action="store_true",
        help="抓取列表全部指數並合併為單一 CSV（列表來源含 count=-1）",
    )
    parser.add_argument(
        "--index-name",
        default=DEFAULT_INDEX_NAME,
        help="單筆模式：在列表比對的指數名稱",
    )
    parser.add_argument(
        "--code",
        default="",
        help="單筆模式：直接指定指數代碼，例如 IX0232",
    )
    parser.add_argument("--start", default="2026/03/20", help="開始日期 YYYY/MM/DD")
    parser.add_argument("--end", default="2026/03/30", help="結束日期 YYYY/MM/DD")
    date_quick = parser.add_mutually_exclusive_group()
    date_quick.add_argument(
        "--today",
        action="store_true",
        help="將起訖日皆設為本機「今天」；與 --start/--end 一併指定時以此為準（不可與 --yesterday 併用）",
    )
    date_quick.add_argument(
        "--yesterday",
        action="store_true",
        help="將起訖日皆設為本機「昨天」，適合測試（當日資料尚未上線時）；與 --start/--end 一併指定時以此為準",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="",
        help="輸出 CSV 路徑（單筆預設 output/<CODE>_history.csv；--all 預設 output/all_indexes_history.csv）",
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="顯示瀏覽器視窗（除錯用）",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="僅處理前 N 支指數（0 表示不限制，用於測試）",
    )
    args = parser.parse_args()

    if args.today:
        today_s = date.today().strftime("%Y/%m/%d")
        args.start = today_s
        args.end = today_s
    elif args.yesterday:
        y = date.today() - timedelta(days=1)
        ys = y.strftime("%Y/%m/%d")
        args.start = ys
        args.end = ys

    lim = args.limit if args.limit > 0 else None

    if args.all:
        out = (
            Path(args.output)
            if args.output
            else Path("output") / "all_indexes_history.csv"
        )
        path = scrape_all_merged_csv(
            args.start,
            args.end,
            out,
            headless=not args.headed,
            limit=lim,
        )
        print(f"已合併儲存：{path.resolve()}")
        return

    code = args.code.strip() or resolve_index_code(args.index_name)
    out = (
        Path(args.output)
        if args.output
        else Path("output") / f"{code}_history.csv"
    )
    path = scrape_history_csv(
        code, args.start, args.end, out, headless=not args.headed
    )
    print(f"已儲存：{path.resolve()}")


if __name__ == "__main__":
    main()
