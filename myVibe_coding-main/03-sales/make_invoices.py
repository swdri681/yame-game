"""매출.xlsx를 기준으로 거래처별 월간 청구서 PDF를 만듭니다."""

from __future__ import annotations

import argparse
import math
import os
import re
import sys
import tempfile
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange


프로그램_폴더 = Path(__file__).resolve().parent
매출_파일 = 프로그램_폴더 / "매출.xlsx"
청구서_파일 = 프로그램_폴더 / "청구서.xlsx"
PDF_폴더 = 프로그램_폴더 / "PDF"

매출_시트명 = "매출"
거래처_시트명 = "거래처"
청구서_시트명 = "청구서"

상세_제목행 = 19
상세_시작행 = 20
기본_상세_마지막행 = 35
기본_합계행 = 36
기본_세액행 = 37
기본_마지막행 = 41
기본_상세행수 = 기본_상세_마지막행 - 상세_시작행 + 1

필수_매출열 = (
    "매출일",
    "거래처코드",
    "거래처명",
    "상품명",
    "매출수량",
    "매출금액",
    "단가",
)
필수_거래처열 = ("거래처코드", "거래처명", "우편번호", "주소")


class 청구서오류(Exception):
    """사용자가 확인할 수 있는 작업 오류입니다."""


@dataclass(frozen=True)
class 거래처정보:
    코드: str
    이름: str
    우편번호: str
    주소: str


@dataclass(frozen=True)
class 매출내역:
    원본행: int
    매출일: datetime
    거래처코드: str
    거래처명: str
    상품명: str
    수량: int | float
    단가: int | float
    매출금액: int | float


@dataclass(frozen=True)
class 생성결과:
    거래처: 거래처정보
    상세건수: int
    공급가액: Decimal
    세액: int
    청구금액: Decimal
    PDF경로: Path


def 명령행_읽기() -> tuple[int, int]:
    parser = argparse.ArgumentParser(
        description="지정한 연월의 거래처별 청구서를 PDF로 만듭니다.",
        epilog=(
            "예: python make_invoices.py 2026 10 / "
            "입력값을 생략하면 2026년 10월 청구서를 만듭니다."
        ),
    )
    parser.add_argument(
        "연도", nargs="?", type=int, default=2026, help="청구서 연도(기본값: 2026)"
    )
    parser.add_argument(
        "월", nargs="?", type=int, default=10, help="청구서 월(기본값: 10)"
    )
    args = parser.parse_args()

    if not 1 <= args.월 <= 12:
        parser.error("월은 1부터 12까지 입력해야 합니다.")
    if not 1900 <= args.연도 <= 9999:
        parser.error("연도는 1900부터 9999까지 입력해야 합니다.")
    return args.연도, args.월


def 다음달_첫날(연도: int, 월: int) -> datetime:
    if 월 == 12:
        return datetime(연도 + 1, 1, 1)
    return datetime(연도, 월 + 1, 1)


def 빈값인가(값: object) -> bool:
    return 값 is None or (isinstance(값, str) and not 값.strip())


def 문자열(값: object) -> str:
    return "" if 값 is None else str(값).strip()


def 코드_문자열(값: object, 항목명: str, 행번호: int) -> str:
    if 빈값인가(값):
        raise 청구서오류(f"{행번호}행의 {항목명}이 비어 있습니다.")
    if isinstance(값, bool):
        raise 청구서오류(f"{행번호}행의 {항목명} 값이 올바르지 않습니다: {값!r}")
    if isinstance(값, int):
        return str(값)
    if isinstance(값, float):
        if not math.isfinite(값) or not 값.is_integer():
            raise 청구서오류(f"{행번호}행의 {항목명} 값이 올바르지 않습니다: {값!r}")
        return str(int(값))
    결과 = str(값).strip()
    if 결과.endswith(".0") and 결과[:-2].isdigit():
        결과 = 결과[:-2]
    return 결과


def 우편번호_문자열(값: object, 셀_표시값: str, 행번호: int) -> str:
    if 빈값인가(값):
        raise 청구서오류(f"거래처 시트 {행번호}행의 우편번호가 비어 있습니다.")
    if isinstance(값, bool):
        raise 청구서오류(f"거래처 시트 {행번호}행의 우편번호가 올바르지 않습니다.")
    if isinstance(값, int):
        결과 = f"{값:05d}"
    elif isinstance(값, float) and math.isfinite(값) and 값.is_integer():
        결과 = f"{int(값):05d}"
    else:
        결과 = str(값).strip()

    # 숫자값이 사용자 지정 표시 형식으로 보이는 경우 표시값을 우선합니다.
    표시값 = str(셀_표시값).strip()
    if 표시값 and 표시값.isdigit() and len(표시값) >= len(결과):
        결과 = 표시값
    return 결과


def 숫자(값: object, 항목명: str, 행번호: int) -> int | float:
    if 빈값인가(값) or isinstance(값, bool):
        raise 청구서오류(f"매출 시트 {행번호}행의 {항목명}이 숫자가 아닙니다: {값!r}")
    if isinstance(값, (int, float)):
        if isinstance(값, float) and not math.isfinite(값):
            raise 청구서오류(f"매출 시트 {행번호}행의 {항목명}이 유효한 숫자가 아닙니다.")
        return 값
    try:
        변환값 = Decimal(str(값).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        raise 청구서오류(
            f"매출 시트 {행번호}행의 {항목명}이 숫자가 아닙니다: {값!r}"
        ) from None
    if not 변환값.is_finite():
        raise 청구서오류(f"매출 시트 {행번호}행의 {항목명}이 유효한 숫자가 아닙니다.")
    return int(변환값) if 변환값 == 변환값.to_integral_value() else float(변환값)


def 소수(값: int | float) -> Decimal:
    return Decimal(str(값))


def 날짜(값: object, 행번호: int) -> datetime:
    if isinstance(값, datetime):
        return 값
    if isinstance(값, date):
        return datetime.combine(값, datetime.min.time())
    raise 청구서오류(
        f"매출 시트 {행번호}행의 매출일이 엑셀 날짜가 아닙니다: {값!r}"
    )


def 제목_위치(시트, 필수열: Iterable[str], 시트명: str) -> dict[str, int]:
    제목들: dict[str, int] = {}
    for 셀 in 시트[1]:
        제목 = 문자열(셀.value)
        if 제목:
            if 제목 in 제목들:
                raise 청구서오류(f"{시트명} 시트 1행에 제목 '{제목}'이 중복되어 있습니다.")
            제목들[제목] = 셀.column

    누락 = [열 for 열 in 필수열 if 열 not in 제목들]
    if 누락:
        raise 청구서오류(f"{시트명} 시트에 필요한 열이 없습니다: {', '.join(누락)}")
    return 제목들


def 입력_파일_확인() -> None:
    없는파일 = [경로 for 경로 in (매출_파일, 청구서_파일) if not 경로.is_file()]
    if 없는파일:
        raise 청구서오류("필요한 파일이 없습니다: " + ", ".join(str(경로) for 경로 in 없는파일))


def 거래처_읽기(통합문서) -> dict[str, 거래처정보]:
    if 거래처_시트명 not in 통합문서.sheetnames:
        raise 청구서오류(f"매출.xlsx에 '{거래처_시트명}' 시트가 없습니다.")
    시트 = 통합문서[거래처_시트명]
    열 = 제목_위치(시트, 필수_거래처열, 거래처_시트명)
    결과: dict[str, 거래처정보] = {}

    for 행번호 in range(2, 시트.max_row + 1):
        행값 = [시트.cell(행번호, 열[제목]).value for 제목 in 필수_거래처열]
        if all(빈값인가(값) for 값 in 행값):
            continue

        코드 = 코드_문자열(시트.cell(행번호, 열["거래처코드"]).value, "거래처코드", 행번호)
        if 코드 in 결과:
            raise 청구서오류(f"거래처 시트에 거래처코드 '{코드}'가 중복되어 있습니다.")

        이름 = 문자열(시트.cell(행번호, 열["거래처명"]).value)
        주소 = 문자열(시트.cell(행번호, 열["주소"]).value)
        우편번호셀 = 시트.cell(행번호, 열["우편번호"])
        우편번호 = 우편번호_문자열(우편번호셀.value, 우편번호셀.value, 행번호)
        누락 = [항목 for 항목, 값 in (("거래처명", 이름), ("우편번호", 우편번호), ("주소", 주소)) if not 값]
        if 누락:
            raise 청구서오류(
                f"거래처 시트 {행번호}행({코드})에 값이 없습니다: {', '.join(누락)}"
            )
        결과[코드] = 거래처정보(코드, 이름, 우편번호, 주소)

    if not 결과:
        raise 청구서오류("거래처 시트에 거래처 데이터가 없습니다.")
    return 결과


def 매출_읽기(통합문서, 연도: int, 월: int, 거래처들: dict[str, 거래처정보]):
    if 매출_시트명 not in 통합문서.sheetnames:
        raise 청구서오류(f"매출.xlsx에 '{매출_시트명}' 시트가 없습니다.")
    시트 = 통합문서[매출_시트명]
    열 = 제목_위치(시트, 필수_매출열, 매출_시트명)
    시작일 = datetime(연도, 월, 1)
    종료일 = 다음달_첫날(연도, 월)
    묶음: dict[str, list[매출내역]] = defaultdict(list)
    전체건수 = 0

    for 행번호 in range(2, 시트.max_row + 1):
        날짜값 = 시트.cell(행번호, 열["매출일"]).value
        관련값 = [시트.cell(행번호, 열[제목]).value for 제목 in 필수_매출열]
        if all(빈값인가(값) for 값 in 관련값):
            continue
        전체건수 += 1
        매출일 = 날짜(날짜값, 행번호)
        if not 시작일 <= 매출일 < 종료일:
            continue

        코드 = 코드_문자열(
            시트.cell(행번호, 열["거래처코드"]).value, "거래처코드", 행번호
        )
        if 코드 not in 거래처들:
            raise 청구서오류(
                f"매출 시트 {행번호}행의 거래처코드 '{코드}'가 거래처 시트에 없습니다."
            )

        거래처명 = 문자열(시트.cell(행번호, 열["거래처명"]).value)
        if not 거래처명:
            raise 청구서오류(f"매출 시트 {행번호}행의 거래처명이 비어 있습니다.")
        if 거래처명 != 거래처들[코드].이름:
            raise 청구서오류(
                f"매출 시트 {행번호}행의 거래처명이 거래처 시트와 다릅니다: "
                f"'{거래처명}' / '{거래처들[코드].이름}'"
            )

        상품명 = 문자열(시트.cell(행번호, 열["상품명"]).value)
        if not 상품명:
            raise 청구서오류(f"매출 시트 {행번호}행의 상품명이 비어 있습니다.")
        수량 = 숫자(시트.cell(행번호, 열["매출수량"]).value, "매출수량", 행번호)
        단가 = 숫자(시트.cell(행번호, 열["단가"]).value, "단가", 행번호)
        매출금액 = 숫자(시트.cell(행번호, 열["매출금액"]).value, "매출금액", 행번호)

        계산금액 = 소수(수량) * 소수(단가)
        if 계산금액 != 소수(매출금액):
            raise 청구서오류(
                f"매출 시트 {행번호}행의 금액이 다릅니다. "
                f"수량×단가={계산금액:,}, 매출금액={소수(매출금액):,}"
            )

        묶음[코드].append(
            매출내역(행번호, 매출일, 코드, 거래처명, 상품명, 수량, 단가, 매출금액)
        )

    if not 묶음:
        raise 청구서오류(f"{연도}년 {월}월 매출이 없습니다. PDF를 만들지 않았습니다.")

    for 내역들 in 묶음.values():
        내역들.sort(key=lambda 항목: (항목.매출일, 항목.원본행))
    return dict(sorted(묶음.items())), 전체건수


def 청구서_양식_확인() -> None:
    통합문서 = load_workbook(청구서_파일, data_only=False, read_only=False)
    try:
        if 청구서_시트명 not in 통합문서.sheetnames:
            raise 청구서오류(f"청구서.xlsx에 '{청구서_시트명}' 시트가 없습니다.")
        시트 = 통합문서[청구서_시트명]
        기대값 = {
            "A5": "우편번호",
            "A6": "주소",
            "A7": "거래처",
            "A19": "날짜",
            "B19": "상품명",
            "F19": "수량",
            "G19": "단가",
            "H19": "금액",
        }
        for 좌표, 값 in 기대값.items():
            if 문자열(시트[좌표].value) != 값:
                raise 청구서오류(
                    f"청구서 양식의 {좌표} 값이 예상과 다릅니다. "
                    f"기대값='{값}', 실제값={시트[좌표].value!r}"
                )
        for 행 in range(상세_시작행, 기본_상세_마지막행 + 1):
            if 문자열(시트[f"H{행}"].value) != f"=F{행}*G{행}":
                raise 청구서오류(f"청구서 양식 H{행}의 금액 수식이 예상과 다릅니다.")
        if 문자열(시트[f"H{기본_합계행}"].value) != "=SUM(H20:H35)":
            raise 청구서오류("청구서 양식의 과세대상액 수식이 예상과 다릅니다.")
        if 문자열(시트[f"H{기본_세액행}"].value) != "=INT(H36*0.1)":
            raise 청구서오류("청구서 양식의 세액 수식이 예상과 다릅니다.")
    finally:
        통합문서.close()


def 병합범위_저장후_해제(시트) -> list[CellRange]:
    기존범위 = [CellRange(str(범위)) for 범위 in 시트.merged_cells.ranges]
    for 범위 in list(시트.merged_cells.ranges):
        시트.unmerge_cells(str(범위))
    return 기존범위


def 병합범위_복원(시트, 기존범위: list[CellRange], 삽입행: int, 이동행수: int) -> None:
    for 범위 in 기존범위:
        if 범위.min_row >= 삽입행:
            범위.shift(row_shift=이동행수)
        시트.merge_cells(str(범위))


def 행_서식_복사(시트, 원본행: int, 대상행: int) -> None:
    for 열번호 in range(1, 시트.max_column + 1):
        원본 = 시트.cell(원본행, 열번호)
        대상 = 시트.cell(대상행, 열번호)
        if 원본.has_style:
            대상._style = copy(원본._style)
        대상.number_format = 원본.number_format
        대상.font = copy(원본.font)
        대상.fill = copy(원본.fill)
        대상.border = copy(원본.border)
        대상.alignment = copy(원본.alignment)
        대상.protection = copy(원본.protection)
    시트.row_dimensions[대상행].height = 시트.row_dimensions[원본행].height
    시트.row_dimensions[대상행].hidden = 시트.row_dimensions[원본행].hidden


def 상세행_확장(시트, 상세건수: int) -> tuple[int, int, int]:
    추가행수 = max(0, 상세건수 - 기본_상세행수)
    if 추가행수:
        기존병합범위 = 병합범위_저장후_해제(시트)
        시트.insert_rows(기본_합계행, amount=추가행수)
        병합범위_복원(시트, 기존병합범위, 기본_합계행, 추가행수)
        for 행 in range(기본_합계행, 기본_합계행 + 추가행수):
            행_서식_복사(시트, 기본_상세_마지막행, 행)
            시트.merge_cells(start_row=행, start_column=2, end_row=행, end_column=5)

    합계행 = 기본_합계행 + 추가행수
    세액행 = 기본_세액행 + 추가행수
    마지막행 = 기본_마지막행 + 추가행수
    return 합계행, 세액행, 마지막행


def 청구서_작업파일_만들기(
    임시경로: Path, 거래처: 거래처정보, 내역들: list[매출내역]
) -> tuple[Decimal, int, Decimal]:
    통합문서 = load_workbook(청구서_파일, data_only=False, read_only=False)
    try:
        시트 = 통합문서[청구서_시트명]
        합계행, 세액행, 마지막행 = 상세행_확장(시트, len(내역들))

        시트["B5"] = 거래처.우편번호
        시트["B6"] = 거래처.주소
        시트["B7"] = 거래처.이름

        # 기본 상세행에 남아 있을 수 있는 값을 먼저 지웁니다.
        실제_상세_마지막행 = max(기본_상세_마지막행, 상세_시작행 + len(내역들) - 1)
        for 행 in range(상세_시작행, 실제_상세_마지막행 + 1):
            for 열 in (1, 2, 6, 7, 9):
                셀 = 시트.cell(행, 열)
                if not isinstance(셀, MergedCell):
                    셀.value = None
            시트.cell(행, 8).value = f"=F{행}*G{행}"

        for 순번, 내역 in enumerate(내역들):
            행 = 상세_시작행 + 순번
            시트.cell(행, 1).value = 내역.매출일
            시트.cell(행, 2).value = 내역.상품명
            시트.cell(행, 6).value = 내역.수량
            시트.cell(행, 7).value = 내역.단가
            시트.cell(행, 8).value = f"=F{행}*G{행}"

        마지막_데이터행 = 상세_시작행 + len(내역들) - 1
        시트[f"H{합계행}"] = f"=SUM(H{상세_시작행}:H{마지막_데이터행})"
        시트[f"H{세액행}"] = f"=INT(H{합계행}*0.1)"
        시트["E17"] = f"=H{합계행}"
        시트["G17"] = f"=H{세액행}"
        시트["I17"] = "=E17+G17"

        시트.print_area = f"A1:I{마지막행}"
        시트.print_title_rows = f"{상세_제목행}:{상세_제목행}"
        시트.page_setup.orientation = "portrait"
        시트.page_setup.fitToWidth = 1
        시트.page_setup.fitToHeight = 0
        시트.sheet_properties.pageSetUpPr.fitToPage = True

        통합문서.calculation.fullCalcOnLoad = True
        통합문서.calculation.forceFullCalc = True
        통합문서.calculation.calcMode = "auto"
        통합문서.save(임시경로)
    finally:
        통합문서.close()

    공급가액 = sum((소수(내역.수량) * 소수(내역.단가) for 내역 in 내역들), Decimal(0))
    세액 = int(공급가액 * Decimal("0.1"))
    return 공급가액, 세액, 공급가액 + 세액


def 안전한_파일명(이름: str) -> str:
    결과 = re.sub(r'[\\/:*?"<>|]', "_", 이름).strip().rstrip(".")
    return 결과 or "이름없음"


def PDF_내보내기(excel, 작업파일: Path, 임시PDF: Path) -> None:
    통합문서 = None
    try:
        통합문서 = excel.Workbooks.Open(str(작업파일), UpdateLinks=0, ReadOnly=True)
        excel.CalculateFullRebuild()
        시트 = 통합문서.Worksheets(청구서_시트명)
        시트.ExportAsFixedFormat(0, str(임시PDF))
    finally:
        if 통합문서 is not None:
            통합문서.Close(SaveChanges=False)


def 청구서들_생성(
    연도: int,
    월: int,
    거래처들: dict[str, 거래처정보],
    매출묶음: dict[str, list[매출내역]],
) -> list[생성결과]:
    try:
        import pythoncom
        import win32com.client
    except ImportError as 오류:
        raise 청구서오류(
            "PDF 생성에 필요한 pywin32(win32com)가 설치되어 있지 않습니다."
        ) from 오류

    PDF_폴더.mkdir(parents=True, exist_ok=True)
    결과들: list[생성결과] = []
    excel = None
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        with tempfile.TemporaryDirectory(prefix="청구서_", dir=프로그램_폴더) as 임시폴더문자열:
            임시폴더 = Path(임시폴더문자열)
            for 코드, 내역들 in 매출묶음.items():
                거래처 = 거래처들[코드]
                작업파일 = 임시폴더 / f"{코드}.xlsx"
                임시PDF = 임시폴더 / f"{코드}.pdf"
                최종이름 = f"{연도:04d}-{월:02d}_청구서_{코드}_{안전한_파일명(거래처.이름)}.pdf"
                최종PDF = PDF_폴더 / 최종이름

                try:
                    공급가액, 세액, 청구금액 = 청구서_작업파일_만들기(
                        작업파일, 거래처, 내역들
                    )
                    PDF_내보내기(excel, 작업파일, 임시PDF)
                except Exception as 오류:
                    if isinstance(오류, 청구서오류):
                        raise
                    raise 청구서오류(
                        f"거래처 {코드}({거래처.이름})의 PDF 생성에 실패했습니다: {오류}"
                    ) from 오류

                if not 임시PDF.is_file() or 임시PDF.stat().st_size == 0:
                    raise 청구서오류(
                        f"거래처 {코드}({거래처.이름})의 PDF가 정상적으로 생성되지 않았습니다."
                    )
                os.replace(임시PDF, 최종PDF)
                결과들.append(
                    생성결과(거래처, len(내역들), 공급가액, 세액, 청구금액, 최종PDF)
                )
    except 청구서오류:
        raise
    except Exception as 오류:
        raise 청구서오류(f"Microsoft Excel을 사용한 PDF 변환에 실패했습니다: {오류}") from 오류
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
    return 결과들


def 금액표시(값: Decimal | int) -> str:
    return f"{값:,.0f}원"


def 실행(연도: int, 월: int) -> None:
    print(f"[1/4] {연도}년 {월}월 청구서 작업을 검사합니다.")
    입력_파일_확인()
    청구서_양식_확인()

    print("[2/4] 거래처와 매출 데이터를 읽고 검산합니다.")
    통합문서 = load_workbook(매출_파일, data_only=False, read_only=False)
    try:
        거래처들 = 거래처_읽기(통합문서)
        매출묶음, 전체건수 = 매출_읽기(통합문서, 연도, 월, 거래처들)
    finally:
        통합문서.close()

    월매출건수 = sum(len(내역들) for 내역들 in 매출묶음.values())

    print(f"[3/4] 거래처 {len(매출묶음)}곳의 PDF를 만듭니다.")
    결과들 = 청구서들_생성(연도, 월, 거래처들, 매출묶음)
    if len(결과들) != len(매출묶음):
        raise 청구서오류("생성한 PDF 수와 거래처 수가 다릅니다.")

    print("[4/4] 생성 결과를 확인했습니다.")
    print()
    print(f"대상 연월: {연도}년 {월}월")
    print(f"전체 매출 데이터: {전체건수:,}건")
    print(f"해당 월 매출: {월매출건수:,}건")
    print(f"청구서 거래처: {len(결과들):,}곳")
    print(f"생성한 PDF: {len(결과들):,}개")
    print()
    for 결과 in 결과들:
        print(
            f"- {결과.거래처.코드} {결과.거래처.이름}: "
            f"{결과.상세건수}건, 공급가액 {금액표시(결과.공급가액)}, "
            f"세액 {금액표시(결과.세액)}, 청구액 {금액표시(결과.청구금액)}"
        )
        print(f"  {결과.PDF경로}")
    print()
    print("주의: 이번 매출에서 제외된 거래처의 기존 같은 달 PDF는 자동으로 삭제하지 않았습니다.")


def main() -> int:
    try:
        연도, 월 = 명령행_읽기()
        실행(연도, 월)
        return 0
    except 청구서오류 as 오류:
        print(f"오류: {오류}", file=sys.stderr)
        return 1
    except PermissionError as 오류:
        print(
            "오류: 파일을 사용할 수 없습니다. 매출.xlsx와 청구서.xlsx를 Excel에서 닫은 뒤 다시 실행하세요. "
            f"({오류})",
            file=sys.stderr,
        )
        return 1
    except Exception as 오류:
        print(f"예상하지 못한 오류가 발생했습니다: {오류}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
