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
    # 部分指數會顯示「查無指數相關歷史資料」浮層，一般 click 會被擋；force 可略過擋點擊檢查
    dates.nth(0).fill(start_date, force=True)
    dates.nth(1).fill(end_date, force=True)

    main = page.locator("main")
    main.get_by_role("button", name=BTN_SEARCH).last.click()
    page.wait_for_timeout(1500)

    with page.expect_download(timeout=60_000) as dl_info:
        main.get_by_role("button", name=BTN_DOWNLOAD).click()
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
    return [row for row in reader if any((v or "").strip() for v in row.values())]


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
