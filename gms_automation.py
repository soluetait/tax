"""
GMS 전표 자동 입력 (Dasan GMS - 세금계산서).

플로우 (1건씩):
  1) 생성 버튼 클릭 → 새 행 추가
  2) 새 행 셀들 순서대로 값 입력
     [1] 증빙일자  [2] 전기일자  [3] 코스트센터(기본값 유지)
     [4] 거래처(사업자번호 + Enter)
     [5] 계정코드(코드 + Enter)
     [6] TX/증빙구분(기본값 유지)
     [7] 공급가액  [8] 기타/봉사료(0)  [9] 부가세  [10] 금액(자동)
     [11] 적요  [12] 전자(기본값 유지)
  3) 저장 버튼 → 확인 다이얼로그 → 예
  4) 증빙 셀 클릭 → 업로드 모달 → 파일 세팅 → 모달 닫기
  5) 다음 건
"""
from __future__ import annotations

import asyncio
import os
from datetime import datetime
from pathlib import Path
from typing import Callable

from playwright.async_api import (
    BrowserContext,
    Page,
    Playwright,
    TimeoutError as PwTimeout,
    async_playwright,
)

GMS_URL = "https://gms.dasannetworks.com/"

APP_DIR = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / "PopbillTaxInvoice"
APP_DIR.mkdir(parents=True, exist_ok=True)
GMS_STATE_FILE = APP_DIR / "gms_state.json"


def _g(d: dict, *keys, default=""):
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return v
    return default


def _fmt_date(s: str) -> str:
    """YYYYMMDD → YYYY-MM-DD."""
    if not s or len(s) != 8:
        return s or ""
    return f"{s[:4]}-{s[4:6]}-{s[6:]}"


def _normalize_biznum(num: str) -> str:
    return "".join(c for c in (num or "") if c.isdigit())


# ============================================================
# GMS 봇
# ============================================================
class GMSBot:
    def __init__(self, log: Callable[[str], None] = print) -> None:
        self.log = log
        self._pw: Playwright | None = None
        self._browser = None
        self._context: BrowserContext | None = None
        self._page: Page | None = None

    async def start(self, headless: bool = False) -> None:
        self._pw = await async_playwright().start()
        try:
            self._browser = await self._pw.chromium.launch(
                headless=headless, channel="msedge"
            )
        except Exception:
            self._browser = await self._pw.chromium.launch(headless=headless)

        state = str(GMS_STATE_FILE) if GMS_STATE_FILE.exists() else None
        self._context = await self._browser.new_context(
            storage_state=state,
            viewport={"width": 1440, "height": 900},
            accept_downloads=True,
        )
        self._page = await self._context.new_page()

    async def close(self) -> None:
        try:
            if self._context:
                await self._context.storage_state(path=str(GMS_STATE_FILE))
        except Exception:
            pass
        try:
            if self._browser:
                await self._browser.close()
        except Exception:
            pass
        try:
            if self._pw:
                await self._pw.stop()
        except Exception:
            pass

    # ---------- 로그인 ----------
    async def login_and_go_taxbill(self, login_timeout_sec: int = 300) -> None:
        assert self._page is not None
        page = self._page
        self.log(f"GMS 접속: {GMS_URL}")
        await page.goto(GMS_URL, wait_until="domcontentloaded", timeout=60000)
        await asyncio.sleep(2)

        # 로그인 버튼 자동 클릭 (세션 없을 때)
        need_login = False
        try:
            btn = page.locator("text=/O365.*로그인/").first
            if await btn.count() > 0 and await btn.is_visible():
                need_login = True
                self.log("O365 로그인 버튼 클릭")
                await btn.click()
        except Exception:
            pass

        if need_login:
            self.log("브라우저에서 O365 계정으로 로그인하세요. (최대 5분 대기)")
            import time as _t
            deadline = _t.time() + login_timeout_sec
            while _t.time() < deadline:
                await asyncio.sleep(2)
                try:
                    url = page.url
                except Exception:
                    continue
                if ("gms.dasannetworks.com" in url
                        and "login" not in url.lower()
                        and "microsoft" not in url.lower()
                        and "auth" not in url.lower()):
                    await asyncio.sleep(3)
                    break
            else:
                raise TimeoutError("GMS 로그인 시간 초과")
            await self._context.storage_state(path=str(GMS_STATE_FILE))  # type: ignore
            self.log(f"세션 저장: {GMS_STATE_FILE.name}")

        # 세금계산서 메뉴로 이동
        await self._navigate_to_taxbill()

    async def _navigate_to_taxbill(self) -> None:
        assert self._page is not None
        page = self._page
        self.log("세금계산서 페이지로 이동 (URL 직접)")

        # 여러 URL 후보 시도 (탐색 덤프에서 확인한 URL 사용)
        candidate_urls = [
            "https://gms.dasannetworks.com/nost/Apply/TaxBill/:type/:code",
            "https://gms.dasannetworks.com/nost/Apply/TaxBill",
        ]
        last_err: Exception | None = None
        for url in candidate_urls:
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=15000)
                # URL 이 세금계산서 페이지로 유지되는지 확인
                await asyncio.sleep(2)
                cur = page.url
                self.log(f"  이동 후 URL: {cur}")
                if "TaxBill" in cur or "tax" in cur.lower():
                    last_err = None
                    break
            except Exception as e:
                last_err = e
                self.log(f"  URL 이동 실패: {url} ({e})")

        # URL 이동으로 안 되면 메뉴 클릭 fallback
        if last_err is not None or "TaxBill" not in page.url:
            self.log("URL 직접 이동 실패, 메뉴 클릭 시도")
            try:
                await page.locator(
                    "a.vsm-link", has_text="전표"
                ).first.click(timeout=5000)
                await asyncio.sleep(1)
            except Exception as e:
                self.log(f"  전표 메뉴 클릭 실패: {e}")
            for sel in [
                'a.vsm-link:has-text("세금계산서")',
                'a:has-text("세금계산서")',
                'text=세금계산서',
            ]:
                try:
                    await page.locator(sel).first.click(timeout=5000)
                    break
                except Exception:
                    continue

        # 페이지 로딩 대기
        try:
            await page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        await asyncio.sleep(2)
        self.log(f"  최종 URL: {page.url}")

        # 세금계산서 페이지 확인
        if "TaxBill" not in page.url and "tax" not in page.url.lower():
            raise RuntimeError(
                f"세금계산서 화면으로 이동 실패. 현재 URL: {page.url}"
            )

    # ---------- 전표 1건 입력 ----------
    async def enter_one_voucher(
        self,
        invoice: dict,
        vendor_info: dict,
        pdf_path: Path | None,
    ) -> None:
        """
        invoice: 팝빌 매입 세금계산서 1건 (dict)
        vendor_info: {"account_code", "account_name", "memo", "vendor_name"}
        pdf_path: 첨부할 PDF 경로 (없으면 첨부 생략)
        """
        assert self._page is not None
        page = self._page

        biznum = _normalize_biznum(_g(invoice, "invoicerCorpNum", "supplierCorpNum"))
        issue_date = _fmt_date(_g(invoice, "issueDate", "writeDate"))
        supply = _g(invoice, "supplyCostTotal", default=0)
        tax = _g(invoice, "taxTotal", default=0)
        item_name = _g(invoice, "itemName", "itemName1")

        # 품목 규칙 해석
        rules = vendor_info.get("rules") or []
        account = vendor_info.get("account") or vendor_info.get("account_code", "")
        cost_center = vendor_info.get("cost_center", "")
        memo_template = vendor_info.get("memo") or ""
        for rule in rules:
            kw = (rule.get("keyword") or "").strip()
            if kw and kw in (item_name or ""):
                account = rule.get("account") or account
                cost_center = rule.get("cost_center") or cost_center
                memo_template = rule.get("memo") or memo_template
                break

        # 적요 플레이스홀더 치환: {월}, {년}, {일}, {month}, {year}, {day}
        memo = memo_template
        if memo and issue_date:
            try:
                y, m, d = issue_date.split("-")
                replacements = {
                    "{월}": f"{int(m)}월",
                    "{년}": f"{y}년",
                    "{일}": f"{int(d)}일",
                    "{month}": str(int(m)),
                    "{year}": y,
                    "{day}": str(int(d)),
                }
                for k, v in replacements.items():
                    memo = memo.replace(k, v)
            except Exception:
                pass

        self.log(f"  [전표] {biznum} {issue_date} 공급가={supply} 세액={tax}")

        # 1. 생성 버튼 클릭
        try:
            btn = page.locator('button:has-text("생성")').first
            await btn.wait_for(state="visible", timeout=8000)
            await btn.click(timeout=5000)
        except Exception as e:
            raise RuntimeError(f"생성 버튼 클릭 실패: {e}")
        await asyncio.sleep(2)

        # 테이블 가로 스크롤을 맨 왼쪽으로 초기화 (cell 1,2,3 접근 가능하게)
        try:
            await page.evaluate("""
                () => {
                    const wrapper = document.querySelector('.v-data-table__wrapper');
                    if (wrapper) wrapper.scrollLeft = 0;
                }
            """)
            await asyncio.sleep(0.3)
        except Exception:
            pass

        # 2. 새로 생성된 행 = v-data-table 의 마지막 행
        # wait_for 는 아래 pick_date 에서 cell 단위로 하므로 여기선 생략
        await asyncio.sleep(0.5)

        # === 셀 인덱스 (v-data-table) ===
        # 0:체크 1:증빙일자 2:전기일자 3:코스트센터 4:거래처 5:계정 6:TX 7:공급가 8:기타 9:부가세 10:금액 11:적요 12:전자

        def last_row():
            return page.locator(".v-data-table tbody tr").last

        async def pick_date(cell_idx: int, iso_date: str) -> None:
            """달력 팝업에서 날짜 선택."""
            import re as _re
            y, m, d = map(int, iso_date.split("-"))
            target_day = d
            target_month = m
            target_year = y

            row = last_row()
            inp = row.locator("td").nth(cell_idx).locator("input:visible").first
            await inp.wait_for(state="attached", timeout=5000)
            await inp.scroll_into_view_if_needed(timeout=5000)
            await inp.click()
            await asyncio.sleep(0.6)

            # 월 네비게이션 (현재 달력과 목표 연/월 비교)
            for _ in range(24):
                try:
                    header = await page.locator(
                        ".v-date-picker-header button:not(.v-btn--icon):visible"
                    ).first.inner_text(timeout=1500)
                except Exception:
                    header = ""
                mobj = _re.match(r"\s*(\d{4})년\s*(\d{1,2})월", header or "")
                if not mobj:
                    break
                cy, cm = int(mobj.group(1)), int(mobj.group(2))
                if cy == target_year and cm == target_month:
                    break
                nav_icons = page.locator(".v-date-picker-header button.v-btn--icon:visible")
                nc = await nav_icons.count()
                if nc < 2:
                    break
                if (cy, cm) < (target_year, target_month):
                    await nav_icons.last.click()  # next
                else:
                    await nav_icons.first.click()  # prev
                await asyncio.sleep(0.3)

            # 해당 일자 클릭
            try:
                await page.locator(
                    f'.v-date-picker-table button:visible >> text=/^{target_day}일$/'
                ).first.click(timeout=3000)
            except Exception as e:
                self.log(f"    {target_day}일 클릭 실패: {e}")
                raise
            await asyncio.sleep(0.3)
            await page.keyboard.press("Escape")
            await asyncio.sleep(0.3)

        async def autocomplete_cell(cell_idx: int, value: str) -> None:
            """거래처/계정: 타이핑 → 드롭다운 → ArrowDown → Enter."""
            row = last_row()
            inp = row.locator("td").nth(cell_idx).locator("input:visible").first
            await inp.wait_for(state="attached", timeout=5000)
            await inp.scroll_into_view_if_needed(timeout=5000)
            await inp.click(click_count=3)
            await inp.press("Delete")
            await inp.type(str(value), delay=30)
            await asyncio.sleep(1.2)
            await page.keyboard.press("ArrowDown")
            await asyncio.sleep(0.3)
            await page.keyboard.press("Enter")
            await asyncio.sleep(0.8)

        async def select_dropdown(cell_idx: int, option_text: str) -> None:
            """TX 같은 v-select 셀: 클릭 → 드롭다운 항목 클릭."""
            row = last_row()
            inp = row.locator("td").nth(cell_idx).locator("input:visible").first
            await inp.wait_for(state="attached", timeout=5000)
            await inp.scroll_into_view_if_needed(timeout=5000)

            # 이미 원하는 값이면 스킵
            try:
                current = (await inp.input_value()).strip()
                if current == option_text:
                    self.log(f"    TX 이미 '{option_text}' - 스킵")
                    return
            except Exception:
                pass

            await inp.click()
            await asyncio.sleep(0.7)
            # visible 한 드롭다운 항목만 선택
            item = page.locator(
                f'.v-menu__content:visible .v-list-item:visible:has-text("{option_text}")'
            ).first
            try:
                if await item.count() > 0:
                    await item.click(timeout=3000)
                else:
                    raise Exception("visible menu item not found")
            except Exception:
                # fallback: 키보드 타이핑 + Enter
                self.log("    드롭다운 fallback (키보드 입력)")
                await inp.type(option_text, delay=30)
                await asyncio.sleep(0.6)
                await page.keyboard.press("Enter")
            await asyncio.sleep(0.8)

        async def fill_number(cell_idx: int, value: str | int) -> None:
            """숫자 셀: click → fill → Tab."""
            row = last_row()
            inp = row.locator("td").nth(cell_idx).locator("input:visible").first
            await inp.wait_for(state="attached", timeout=5000)
            await inp.scroll_into_view_if_needed(timeout=5000)
            await inp.click()
            await asyncio.sleep(0.2)
            await inp.fill(str(value))
            await asyncio.sleep(0.3)
            await page.keyboard.press("Tab")
            await asyncio.sleep(0.5)

        async def fill_text(cell_idx: int, value: str) -> None:
            """일반 텍스트 셀 (적요 등)."""
            row = last_row()
            inp = row.locator("td").nth(cell_idx).locator("input:visible").first
            await inp.wait_for(state="attached", timeout=5000)
            await inp.scroll_into_view_if_needed(timeout=5000)
            await inp.click(click_count=3)
            await inp.press("Delete")
            await inp.type(str(value), delay=20)
            await asyncio.sleep(0.3)
            await page.keyboard.press("Tab")
            await asyncio.sleep(0.3)

        # ------ 실제 입력 ------
        if issue_date:
            await pick_date(1, issue_date)
            await pick_date(2, issue_date)

        # 코스트센터 (매핑 있을 때만)
        if cost_center:
            await autocomplete_cell(3, cost_center)

        # 거래처 (사업자번호)
        await autocomplete_cell(4, biznum)

        # 계정
        await autocomplete_cell(5, account)

        # TX = 영세율이면 "영세율", 그 외 "세금계산서" (필수 — 부가세 unlock)
        tax_type = _g(invoice, "taxType")
        if tax_type == "Z":
            await select_dropdown(6, "영세율")
        else:
            await select_dropdown(6, "세금계산서")

        # 공급가액
        supply_val = int(float(str(supply).replace(",", "") or 0))
        await fill_number(7, supply_val)

        # 부가세: 영세율이면 0 (입력 생략), 그 외에는 세액 입력
        tax_val = int(float(str(tax).replace(",", "") or 0))
        if tax_type == "Z":
            pass  # 영세율은 부가세 없음
        elif tax_val != 0:
            await fill_number(9, tax_val)

        # 적요
        if memo:
            await fill_text(11, memo)

        # 3. 저장 버튼 클릭
        self.log("  저장 버튼 클릭")
        await page.locator('button:has-text("저장")').first.click()
        await asyncio.sleep(1)

        # 4. 확인 다이얼로그에서 '저장' 클릭
        # 주의: 데이터테이블의 저장 버튼과 다이얼로그의 저장 버튼이 동명
        # 다이얼로그 내부에서만 찾기
        confirm_clicked = False
        for label in ("저장", "예", "확인", "Yes", "OK"):
            try:
                btn = page.locator(
                    f'.v-dialog--active button:has-text("{label}"), '
                    f'.v-dialog[style*="display"] button:has-text("{label}")'
                ).first
                if await btn.count() > 0 and await btn.is_visible():
                    await btn.click(timeout=3000)
                    self.log(f"  다이얼로그 '{label}' 클릭")
                    confirm_clicked = True
                    break
            except Exception:
                continue
        if not confirm_clicked:
            self.log("  확인 다이얼로그 버튼을 찾지 못함")
        await asyncio.sleep(2.5)

        # 4. 저장 완료 성공 팝업(SweetAlert2) 닫기
        try:
            swal_btn = page.locator(
                '.swal2-container .swal2-confirm, .swal2-container button'
            ).first
            if await swal_btn.count() > 0 and await swal_btn.is_visible():
                await swal_btn.click(timeout=3000)
                self.log("  저장 성공 팝업 닫기")
            else:
                await page.keyboard.press("Enter")
        except Exception:
            try:
                await page.keyboard.press("Enter")
            except Exception:
                pass
        await asyncio.sleep(1)

        # 5. PDF 첨부 (저장 후 행이 증빙일자 순 재정렬되므로 매칭으로 찾기)
        if pdf_path and pdf_path.exists():
            await self._attach_pdf_matched(issue_date, supply_val, pdf_path)
        elif pdf_path:
            self.log(f"  PDF 파일 없음: {pdf_path}")

    async def _attach_pdf_matched(
        self, issue_date: str, supply_val: int, pdf_path: Path
    ) -> None:
        """저장된 전표 중 증빙일자 + 공급가액이 일치하는 행을 찾아 첨부."""
        assert self._page is not None
        page = self._page
        self.log(f"  첨부 시작: {pdf_path.name} (매칭 {issue_date} / {supply_val:,})")

        # 테이블 가로 스크롤 리셋
        try:
            await page.evaluate("""
                () => {
                    const w = document.querySelector('.v-data-table__wrapper');
                    if (w) w.scrollLeft = 0;
                }
            """)
            await asyncio.sleep(0.3)
        except Exception:
            pass

        # 매칭되는 행 찾기: 증빙일자 + 공급가액 일치 AND 증빙 개수 0
        target_row = None
        rows = page.locator(".v-data-table tbody tr")
        n = await rows.count()
        for i in range(n):
            r = rows.nth(i)
            try:
                date_inp = r.locator("td").nth(1).locator("input:visible").first
                if await date_inp.count() == 0:
                    continue
                date_val = (await date_inp.input_value()).strip()
                if date_val != issue_date:
                    continue
                # 공급가액 매칭
                await r.locator("td").nth(7).scroll_into_view_if_needed(timeout=2000)
                supply_inp = r.locator("td").nth(7).locator("input:visible").first
                if await supply_inp.count() == 0:
                    continue
                supply_str = (await supply_inp.input_value()).replace(",", "").strip()
                try:
                    if int(supply_str or 0) != supply_val:
                        continue
                except ValueError:
                    continue
                # 증빙 개수 확인 (cell 14 의 버튼 텍스트)
                await r.locator("td").nth(14).scroll_into_view_if_needed(timeout=2000)
                cell14_text = (await r.locator("td").nth(14).inner_text()).strip()
                # "0" 이면 미첨부, 숫자가 크면 이미 첨부됨
                try:
                    attached_count = int(cell14_text)
                except ValueError:
                    attached_count = 0
                if attached_count > 0:
                    self.log(f"    row {i}: 매칭되지만 이미 증빙 {attached_count}개 존재, 스킵")
                    continue
                target_row = r
                self.log(f"    매칭 행 index={i} (증빙 0개)")
                break
            except Exception:
                continue

        if target_row is None:
            self.log(f"  매칭 행 없음 (모두 이미 첨부됐거나 데이터 불일치)")
            return

        # 증빙 셀(14) 클릭
        try:
            cell14 = target_row.locator("td").nth(14)
            await cell14.scroll_into_view_if_needed(timeout=5000)
            await cell14.click(timeout=5000)
        except Exception as e:
            self.log(f"  증빙 셀 클릭 실패: {e}")
            return

        # Dropzone hidden input 대기
        try:
            await page.wait_for_selector(
                "input.dz-hidden-input", state="attached", timeout=8000
            )
        except PwTimeout:
            self.log("  업로드 모달 입력 대기 실패")
            return

        # 파일 설정 (hidden input 에 직접 set — 파일 대화상자 우회)
        file_input = page.locator("input.dz-hidden-input").first
        await file_input.set_input_files(str(pdf_path))
        self.log("  파일 설정 완료, 업로드 대기")
        await asyncio.sleep(4)

        # 모달 닫기: 헤더의 X 아이콘 (<i class="v-icon v-icon--link ... clear">)
        try:
            x_btn = page.locator(
                '.v-dialog--active i.v-icon--link:has-text("clear")'
            ).first
            await x_btn.click(timeout=4000)
            self.log("  X 버튼 클릭")
        except Exception as e:
            self.log(f"  X 버튼 클릭 실패: {e}")
        await asyncio.sleep(1)


# ============================================================
# 엔트리 포인트 (외부 호출용)
# ============================================================
async def enter_vouchers(
    items: list[dict],
    vendor_mapping: dict,
    download_dir: Path,
    log: Callable[[str], None] = print,
    progress: Callable[[int, int, str], None] | None = None,
    headless: bool = True,
    keep_open: bool = False,
) -> dict:
    """
    headless: True 면 브라우저 창 안 뜨고 백그라운드 실행 (PC 사용 가능)
    keep_open: True 면 완료 후에도 GMS 창 유지 (사용자가 결과 확인 / 결재 상신)
    return: {"ok", "fail", "skipped", "missing_vendors"}
    """
    ok = 0
    fail = 0
    skipped: list[dict] = []
    missing_vendors: list[dict] = []
    success_nts: list[str] = []

    # 세션 없으면 강제 headful (로그인 필요)
    use_headless = headless and GMS_STATE_FILE.exists()
    mode = "백그라운드(headless)" if use_headless else "창 표시(headful)"
    log(f"GMS 브라우저 모드: {mode}")

    bot = GMSBot(log=log)
    await bot.start(headless=use_headless)
    try:
        try:
            await bot.login_and_go_taxbill()
        except (TimeoutError, RuntimeError) as e:
            # 세션 만료 또는 로그인 페이지로 리디렉트된 경우
            if use_headless:
                log(f"세션 만료 감지 ({e}), headful 재시작")
                await bot.close()
                bot = GMSBot(log=log)
                await bot.start(headless=False)
                await bot.login_and_go_taxbill()
            else:
                raise

        total = len(items)
        for i, inv in enumerate(items, 1):
            nts = _g(inv, "ntsconfirmNum", "ntsConfirmNum")
            if progress:
                progress(i, total, nts)

            biznum = _normalize_biznum(_g(inv, "invoicerCorpNum", "supplierCorpNum"))
            info = vendor_mapping.get(biznum)
            account = (info or {}).get("account") or (info or {}).get("account_code", "")
            if not info or not account:
                log(f"  [{i}/{total}] {nts} SKIP - 거래처 매핑 없음 ({biznum})")
                missing_vendors.append({
                    "biznum": biznum,
                    "vendor_name": _g(inv, "invoicerCorpName", "supplierCorpName"),
                    "nts": nts,
                })
                skipped.append(inv)
                continue

            pdf_path = download_dir / f"{nts}.pdf" if nts else None
            try:
                await bot.enter_one_voucher(inv, info, pdf_path)
                ok += 1
                if nts:
                    success_nts.append(nts)
            except Exception as e:
                log(f"  [{i}/{total}] {nts} 실패: {e}")
                fail += 1
    finally:
        if keep_open and not use_headless:
            # headful 모드에서 keep_open 이면 세션만 저장하고 창은 유지
            try:
                if bot._context:
                    await bot._context.storage_state(path=str(GMS_STATE_FILE))
                log("GMS 창 유지 중 (수동으로 닫으세요)")
            except Exception:
                pass
        else:
            await bot.close()

    return {
        "ok": ok,
        "fail": fail,
        "skipped": skipped,
        "missing_vendors": missing_vendors,
        "success_nts": success_nts,
    }
