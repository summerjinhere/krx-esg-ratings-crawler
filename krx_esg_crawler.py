# -*- coding: utf-8 -*-
"""
KRX ESG 포털 크롤러 (로컬 PC용) - Refactored v2 (Robust Click + Debug)
=====================================================================

이번 v2는 사용자 리포트(조회 버튼 클릭 타임아웃 / element is outside of the viewport 반복)를
직접 해결하기 위한 “UI 액션 안정화 + HTML/JSON 구조 확인용 디버그 덤프”가 추가된 버전입니다.

핵심 개선
1) 조회 버튼 클릭을 locator 기반으로 재작성 (scroll_into_view_if_needed → force click → JS click → Enter 키 fallback)
2) 실패 시 페이지 HTML/스크린샷/네트워크(XHR/JSON) 로그를 남겨서 구조 확인 가능
3) --input에 파일/폴더 모두 지원(기존 refactored 유지)

필수 설치:
  pip install playwright pandas openpyxl tqdm
  playwright install chromium

실행 예시(Windows):
  python krx_esg_crawler.py ^
    --input "kcgs_ratings.xlsx" ^
    --sheet "평가등급" ^
    --headless ^
    --debug

디버그 산출물(동일 폴더):
  debug_dump/\n
    - page_*.html
    - page_*.png
    - button_outerhtml_*.txt
    - network_jsonl_*.jsonl
"""

import asyncio
import pandas as pd
import re
import json
import random
import argparse
from datetime import datetime
from pathlib import Path
from typing import Optional
from tqdm import tqdm

try:
    from playwright.async_api import async_playwright
except ImportError:
    print("Playwright가 설치되지 않았습니다.")
    print("설치: pip install playwright && playwright install chromium")
    raise


def _normalize_company_code(x) -> str | None:
    """Excel 기업코드를 6자리 문자열로 정규화."""
    if x is None or (isinstance(x, float) and pd.isna(x)) or (isinstance(x, str) and x.strip() == ""):
        return None
    s = str(x).strip()
    s = re.sub(r"\.0+$", "", s)  # 5930.0 같은 케이스
    digits = re.findall(r"\d+", s)
    if not digits:
        return None
    return "".join(digits).zfill(6)


def _pick_input_xlsx(input_path: str, prefer_keywords: list[str] | None = None) -> Path:
    """파일/폴더 입력을 모두 지원."""
    p = Path(input_path).expanduser()
    if p.is_file():
        return p
    if not p.exists() or not p.is_dir():
        raise FileNotFoundError(f"입력 경로를 찾을 수 없습니다: {p}")

    prefer_keywords = prefer_keywords or ["평가등급", "ESG", "KCGS", "등급"]
    candidates: list[Path] = []
    for ext in ("*.xlsx", "*.xlsm", "*.xls"):
        candidates.extend(sorted(p.glob(ext)))
    if not candidates:
        raise FileNotFoundError(f"폴더 내 Excel 파일을 찾지 못했습니다: {p}")

    def score(fp: Path) -> int:
        name = fp.name.lower()
        sc = 0
        for kw in prefer_keywords:
            if kw.lower() in name:
                sc += 10
        try:
            sc += int(fp.stat().st_mtime // 1_000_000)
        except Exception:
            pass
        return sc

    candidates.sort(key=score, reverse=True)
    return candidates[0]


def load_company_list_from_xlsx(xlsx_path: str, sheet: str | int | None = None) -> pd.DataFrame:
    """XLSX에서 기업명/기업코드 로드."""
    xlsx_path = str(xlsx_path)
    print(f"기업 목록 로드: {xlsx_path}")
    df = pd.read_excel(xlsx_path, sheet_name=sheet if sheet is not None else 0)
    df.columns = [str(c).strip() for c in df.columns]

    required = ["기업명", "기업코드"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}\n현재 컬럼: {df.columns.tolist()}")

    companies = df[["기업명", "기업코드"]].copy()
    companies["기업명"] = companies["기업명"].astype(str).str.strip()
    companies["기업코드"] = companies["기업코드"].apply(_normalize_company_code)
    companies = companies.dropna(subset=["기업코드"])
    companies = companies[companies["기업코드"].str.len() == 6]
    companies = companies.drop_duplicates(subset=["기업코드"], keep="first").reset_index(drop=True)
    print(f"총 {len(companies)}개 기업(중복 제거 후)")
    return companies


class KRXESGCrawler:
    BASE_URL = "https://esg.krx.co.kr/contents/02/02010000/ESG02010000.jsp"
    DEFAULT_YEARS = ["2025", "2024", "2023", "2022", "2021", "2020", "2019"]

    def __init__(
        self,
        headless: bool = True,
        slow_mo: int = 150,
        years: Optional[list[str]] = None,
        debug: bool = False,
    ):
        self.headless = headless
        self.slow_mo = slow_mo
        self.years = years or self.DEFAULT_YEARS
        self.debug = debug

        self.data: list[dict] = []
        self.failed: list[dict] = []
        self.checkpoint_file = "krx_checkpoint.json"

        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

        self._debug_dir = Path("debug_dump")
        self._network_log_path = None

    # ------------------------
    # Debug helpers
    # ------------------------
    def _ts(self) -> str:
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    async def _dump_debug_artifacts(self, tag: str, btn_locator=None):
        if not self.debug or not self.page:
            return

        self._debug_dir.mkdir(parents=True, exist_ok=True)
        ts = self._ts()

        # HTML dump
        html_path = self._debug_dir / f"page_{tag}_{ts}.html"
        try:
            html = await self.page.content()
            html_path.write_text(html, encoding="utf-8")
        except Exception:
            pass

        # Screenshot
        png_path = self._debug_dir / f"page_{tag}_{ts}.png"
        try:
            await self.page.screenshot(path=str(png_path), full_page=True)
        except Exception:
            pass

        # Button outerHTML
        if btn_locator is not None:
            outer_path = self._debug_dir / f"button_outerhtml_{tag}_{ts}.txt"
            try:
                outer = await btn_locator.evaluate("el => el.outerHTML")
                outer_path.write_text(str(outer), encoding="utf-8")
            except Exception:
                pass

    async def _log_json_response(self, resp):
        """XHR/Fetch JSON response를 jsonl로 기록(구조 확인용)."""
        if not self.debug:
            return

        try:
            ct = (resp.headers or {}).get("content-type", "")
            if "application/json" not in ct and "json" not in ct:
                return

            url = resp.url
            # JSON이 너무 크면 로깅이 과해질 수 있어 제한
            txt = await resp.text()
            if len(txt) > 2_000_000:
                txt = txt[:2_000_000] + "\n...TRUNCATED..."

            rec = {
                "ts": datetime.now().isoformat(),
                "url": url,
                "status": resp.status,
                "content_type": ct,
                "body": txt,
            }

            if self._network_log_path:
                with open(self._network_log_path, "a", encoding="utf-8") as f:
                    f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        except Exception:
            return

    # ------------------------
    # Browser lifecycle
    # ------------------------
    async def init_browser(self):
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(
            headless=self.headless,
            slow_mo=self.slow_mo,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-sandbox",
            ],
        )

        self.context = await self.browser.new_context(
            viewport={"width": 1920, "height": 1400},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            locale="ko-KR",
            timezone_id="Asia/Seoul",
        )

        await self.context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
        )

        self.page = await self.context.new_page()
        self.page.set_default_timeout(30_000)

        if self.debug:
            self._debug_dir.mkdir(parents=True, exist_ok=True)
            self._network_log_path = str(self._debug_dir / f"network_jsonl_{self._ts()}.jsonl")
            self.page.on("response", lambda r: asyncio.create_task(self._log_json_response(r)))

    async def close_browser(self):
        try:
            if self.browser:
                await self.browser.close()
        finally:
            if self.playwright:
                await self.playwright.stop()

    async def random_delay(self, min_sec=0.6, max_sec=1.6):
        await asyncio.sleep(random.uniform(min_sec, max_sec))

    # ------------------------
    # Core actions
    # ------------------------
    async def load_main_page(self) -> bool:
        print("KRX ESG 포털 접속 중...")
        try:
            await self.page.goto(self.BASE_URL, wait_until="networkidle", timeout=60_000)
            await self.random_delay(1.5, 2.5)
            content = await self.page.content()
            if "Access Denied" in content or "일시 제한" in content:
                print("❌ IP가 차단되었습니다! (Access Denied / 일시 제한)")
                return False
            print("✓ 페이지 로드 완료")
            return True
        except Exception as e:
            print(f"❌ 페이지 로드 실패: {e}")
            return False

    async def search_company(self, company_name: str) -> bool:
        try:
            search = self.page.locator("input#schCmpyNm, input[name='schCmpyNm']").first
            if await search.count() == 0:
                # fallback: 첫 text input
                search = self.page.locator("input[type='text']").first

            await search.wait_for(state="visible", timeout=10_000)
            await search.click()
            await search.fill("")
            await self.random_delay(0.2, 0.4)
            await search.fill(company_name)
            await self.random_delay(0.4, 0.8)
            return True
        except Exception as e:
            print(f"  검색창 입력 오류: {e}")
            await self._dump_debug_artifacts("search_input_fail")
            return False

    async def select_year(self, year: str) -> bool:
        try:
            # 페이지 내 select 중 year option이 있는 첫 select
            selects = self.page.locator("select")
            for i in range(await selects.count()):
                sel = selects.nth(i)
                try:
                    opt = sel.locator(f"option[value='{year}']")
                    if await opt.count() > 0:
                        await sel.select_option(value=year)
                        await self.random_delay(0.15, 0.35)
                        return True
                except Exception:
                    continue
        except Exception as e:
            print(f"  연도 선택 오류: {e}")
        return False

    async def _safe_click_locator(self, loc, tag: str) -> bool:
        """Playwright actionability 이슈(밖에 있음/가려짐) 대응을 한 번에 처리."""
        try:
            await loc.wait_for(state="visible", timeout=10_000)
        except Exception:
            await self._dump_debug_artifacts(f"{tag}_not_visible")
            return False

        # 1) 표준 클릭 시도
        try:
            await loc.scroll_into_view_if_needed(timeout=5_000)
        except Exception:
            pass

        try:
            await loc.click(timeout=8_000)
            return True
        except Exception:
            pass

        # 2) force click
        try:
            await loc.click(timeout=8_000, force=True)
            return True
        except Exception:
            pass

        # 3) JS click (viewport 제약 우회)
        try:
            await loc.evaluate(
                """
                (el) => {
                  try { el.scrollIntoView({block:'center', inline:'center'}); } catch(e) {}
                  el.click();
                }
                """
            )
            await self.random_delay(0.4, 0.8)
            return True
        except Exception:
            await self._dump_debug_artifacts(f"{tag}_click_fail", btn_locator=loc)
            return False

    async def click_search(self) -> bool:
        """조회 버튼 클릭 (실패 시 HTML/스크린샷/outerHTML 저장 + Enter fallback)."""
        # 가장 “KRX 구조 변화”에 덜 취약한 순서로 시도
        candidates = [
            self.page.locator("#btnSearch"),
            self.page.locator("button:has-text('조회')"),
            self.page.locator("a:has-text('조회')"),
            self.page.locator("input[value='조회']"),
        ]

        for i, loc in enumerate(candidates):
            if await loc.count() == 0:
                continue
            ok = await self._safe_click_locator(loc.first, tag=f"searchbtn_{i}")
            if ok:
                await self.random_delay(1.2, 2.0)
                return True

        # 마지막 fallback: 검색창에서 Enter
        try:
            search = self.page.locator("input#schCmpyNm, input[name='schCmpyNm']").first
            if await search.count() == 0:
                search = self.page.locator("input[type='text']").first
            await search.press("Enter")
            await self.random_delay(1.2, 2.0)
            return True
        except Exception:
            await self._dump_debug_artifacts("search_enter_fail")
            return False

    async def extract_ratings(self, company_name: str, company_code: str, year: str) -> Optional[dict]:
        """(기존 버전과 동일: HTML에서 정규식으로 등급 추출)"""
        data = {
            "기업명": company_name,
            "기업코드": company_code,
            "연도": year,
            "KCGS_ESG": None,
            "KCGS_E": None,
            "KCGS_S": None,
            "KCGS_G": None,
            "서스틴베스트_ESG": None,
            "한국ESG연구소_ESG": None,
            "MSCI_ESG": None,
            "SP_Score": None,
        }

        try:
            content = await self.page.content()
            if "데이터가 없습니다" in content or "검색 결과가 없습니다" in content:
                return None

            # KCGS 통합등급
            kcgs_patterns = [
                r"KCGS[^<]*?([SABCD]\+?)",
                r"통합등급[^<]*?([SABCD]\+?)",
                r"class=\"grade[^\"]*\"[^>]*>([SABCD]\+?)<",
            ]
            for pat in kcgs_patterns:
                m = re.search(pat, content, re.IGNORECASE)
                if m:
                    data["KCGS_ESG"] = m.group(1).upper()
                    break

            # E/S/G
            esg_ind = re.findall(r"([환경사회지배구조])[^<]*?([SABCD]\+?)", content)
            for k, v in esg_ind:
                if k == "환경":
                    data["KCGS_E"] = v
                elif k == "사회":
                    data["KCGS_S"] = v
                elif k == "지배구조":
                    data["KCGS_G"] = v

            # 서스틴베스트
            m = re.search(r"서스틴[^<]*?([A]{1,2}|BB|B|C|D|E)", content)
            if m:
                data["서스틴베스트_ESG"] = m.group(1)

            # 한국ESG연구소
            m = re.search(r"한국ESG[^<]*?([SABCD]\+?)", content)
            if m:
                data["한국ESG연구소_ESG"] = m.group(1)

            # MSCI
            m = re.search(r"MSCI[^<]*?([A]{1,3}|BBB|BB|B|CCC)", content)
            if m:
                data["MSCI_ESG"] = m.group(1)

            # S&P
            m = re.search(r"S&P[^<]*?(\d{1,3})", content)
            if m:
                score = int(m.group(1))
                if 0 <= score <= 100:
                    data["SP_Score"] = score

            if any(v is not None for k, v in data.items() if k not in ("기업명", "기업코드", "연도")):
                return data

        except Exception as e:
            print(f"  등급 추출 오류: {e}")
            await self._dump_debug_artifacts("extract_fail")
        return None

    async def crawl_company(self, company_name: str, company_code: str) -> list[dict]:
        out: list[dict] = []
        if not await self.search_company(company_name):
            return out

        for year in self.years:
            try:
                await self.select_year(year)
                if not await self.click_search():
                    # 조회가 안 되면 다음 year도 의미가 없어서 break
                    return out
                r = await self.extract_ratings(company_name, company_code, year)
                if r:
                    out.append(r)
            except Exception:
                pass
            await self.random_delay(0.8, 1.5)
        return out

    # ------------------------
    # Checkpoint
    # ------------------------
    def save_checkpoint(self, processed_idx: int):
        checkpoint = {
            "processed_idx": processed_idx,
            "data": self.data,
            "failed": self.failed,
            "timestamp": datetime.now().isoformat(),
        }
        with open(self.checkpoint_file, "w", encoding="utf-8") as f:
            json.dump(checkpoint, f, ensure_ascii=False, indent=2)

    def load_checkpoint(self) -> int:
        p = Path(self.checkpoint_file)
        if p.exists():
            with open(p, "r", encoding="utf-8") as f:
                ck = json.load(f)
            self.data = ck.get("data", [])
            self.failed = ck.get("failed", [])
            return int(ck.get("processed_idx", 0))
        return 0

    async def crawl_all(self, companies_df: pd.DataFrame, max_companies: int | None = None, resume: bool = False) -> list[dict]:
        await self.init_browser()

        if not await self.load_main_page():
            await self.close_browser()
            return []

        start_idx = 0
        if resume:
            start_idx = self.load_checkpoint()
            if start_idx > 0:
                print(f"체크포인트에서 재개: {start_idx}번째부터")

        companies = companies_df.iloc[start_idx:].copy()
        if max_companies:
            companies = companies.head(max_companies)

        total = len(companies)
        print(f"\n총 {total}개 기업 크롤링 시작...")
        print("=" * 60)

        for idx, (_, row) in enumerate(tqdm(companies.iterrows(), total=total, desc="크롤링")):
            name = str(row["기업명"]).strip()
            code = str(row["기업코드"]).zfill(6)
            try:
                rows = await self.crawl_company(name, code)
                if rows:
                    self.data.extend(rows)
                else:
                    self.failed.append({"기업명": name, "기업코드": code, "사유": "데이터 없음/조회 실패"})
            except Exception as e:
                self.failed.append({"기업명": name, "기업코드": code, "사유": str(e)})

            if (idx + 1) % 10 == 0:
                self.save_checkpoint(start_idx + idx + 1)
                print(f"\n  체크포인트 저장: {start_idx + idx + 1}개 완료, 누적 {len(self.data)}건")

            await self.random_delay(1.2, 2.4)
            if (idx + 1) % 50 == 0:
                print("\n  잠시 대기 중... (차단 방지)")
                await asyncio.sleep(random.uniform(25, 55))

        await self.close_browser()
        print("\n크롤링 완료!")
        print(f"  성공: {len(self.data)}건")
        print(f"  실패: {len(self.failed)}건")
        return self.data


def _parse_years(years_arg: str | None) -> list[str] | None:
    if not years_arg:
        return None
    years = []
    for y in re.split(r"[,\s]+", years_arg.strip()):
        if re.fullmatch(r"\d{4}", y or ""):
            years.append(y)
    return years or None


async def main():
    parser = argparse.ArgumentParser(description="KRX ESG 포털 크롤러 (XLSX 기업목록 기반) - v2")
    parser.add_argument("--input", "-i", type=str, required=True, help="입력 XLSX 파일 경로 또는 폴더")
    parser.add_argument("--sheet", type=str, default=None, help='기업목록 시트명(예: "평가등급"). 미지정 시 첫 시트')
    parser.add_argument("--output", "-o", type=str, default=None, help="출력 CSV 파일명")
    parser.add_argument("--max", "-m", type=int, default=None, help="최대 크롤링 기업 수")
    parser.add_argument("--headless", action="store_true", help="헤드리스 모드")
    parser.add_argument("--resume", action="store_true", help="체크포인트에서 재개")
    parser.add_argument("--slow-mo", type=int, default=150, help="slow_mo(ms). 차단/오류 시 200~600 권장")
    parser.add_argument("--years", type=str, default=None, help='예: "2019,2020,2021"')
    parser.add_argument("--debug", action="store_true", help="실패 시 HTML/PNG/JSON 로그 저장")
    args = parser.parse_args()

    xlsx = _pick_input_xlsx(args.input)
    companies = load_company_list_from_xlsx(str(xlsx), sheet=args.sheet)

    years = _parse_years(args.years)
    crawler = KRXESGCrawler(
        headless=args.headless,
        slow_mo=int(args.slow_mo),
        years=years,
        debug=bool(args.debug),
    )

    await crawler.crawl_all(companies, max_companies=args.max, resume=args.resume)

    out = args.output
    if not out:
        out = f"krx_esg_ratings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

    df = pd.DataFrame(crawler.data)
    if not df.empty:
        df.to_csv(out, index=False, encoding="utf-8-sig")
        print(f"\n[저장 완료] {out}")
    else:
        print("\n[경고] 저장할 데이터가 없습니다.")

    if crawler.failed:
        failed_path = Path(out).with_name(Path(out).stem + "_failed.csv")
        pd.DataFrame(crawler.failed).to_csv(failed_path, index=False, encoding="utf-8-sig")
        print(f"[실패 목록 저장] {failed_path}")


if __name__ == "__main__":
    asyncio.run(main())
