"""CSV 매출 자료를 검사한 뒤 매출.xlsx에 통합하는 프로그램입니다."""

from __future__ import annotations

import csv
import os
import shutil
import sys
import tempfile
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell import Cell


시트_이름 = "매출"
엑셀_파일명 = "매출.xlsx"
예상_제목 = [
    "매출일",
    "영업소",
    "영업담당",
    "기간",
    "전표번호",
    "상품코드",
    "상품명",
    "대분류",
    "중분류",
    "소분류",
    "거래처코드",
    "거래처명",
    "매출수량",
    "매출금액",
    "매출이익",
    "단가",
]

# 코드 열은 숫자로 바꾸면 072745 같은 값의 첫 번째 0이 사라지므로 문자로 둡니다.
문자_열 = {1, 2, 4, 5, 6, 7, 8, 9, 10, 11}
정수_열 = {3, 12, 13, 14, 15}


class 작업오류(Exception):
    """사용자에게 원인을 설명하고 정상적으로 종료할 수 있는 오류입니다."""


def 작업_경로_만들기() -> tuple[Path, Path, Path]:
    """프로그램 파일의 위치를 기준으로 작업 경로를 만듭니다."""
    기준_폴더 = Path(__file__).resolve().parent
    return (
        기준_폴더 / 엑셀_파일명,
        기준_폴더 / "csv",
        기준_폴더 / "completed",
    )


def csv_목록_찾기(csv_폴더: Path) -> list[Path]:
    """csv 폴더 바로 아래의 CSV 파일만 이름순으로 반환합니다."""
    if not csv_폴더.is_dir():
        raise 작업오류(f"CSV 폴더가 없습니다: {csv_폴더}")

    파일들 = sorted(
        (경로 for 경로 in csv_폴더.iterdir() if 경로.is_file() and 경로.suffix.lower() == ".csv"),
        key=lambda 경로: 경로.name,
    )
    if not 파일들:
        raise 작업오류(f"처리할 CSV 파일이 없습니다: {csv_폴더}")
    return 파일들


def 엑셀_제목_검사(엑셀_경로: Path) -> list[str]:
    """엑셀 파일과 매출 시트를 읽기 전용으로 검사합니다."""
    if not 엑셀_경로.is_file():
        raise 작업오류(f"엑셀 파일이 없습니다: {엑셀_경로}")

    try:
        통합문서 = load_workbook(엑셀_경로, read_only=True, data_only=False)
    except Exception as 오류:
        raise 작업오류(f"엑셀 파일을 열 수 없습니다: {오류}") from 오류

    try:
        if 시트_이름 not in 통합문서.sheetnames:
            raise 작업오류(f"엑셀 파일에 '{시트_이름}' 시트가 없습니다.")
        시트 = 통합문서[시트_이름]
        제목 = [시트.cell(1, 열).value for 열 in range(1, len(예상_제목) + 1)]
        if 제목 != 예상_제목:
            raise 작업오류(
                "엑셀 1행의 제목 또는 열 순서가 예상과 다릅니다.\n"
                f"예상: {예상_제목}\n실제: {제목}"
            )
        return 제목
    finally:
        통합문서.close()


def 정수로_변환(값: str, 파일명: str, 행번호: int, 열이름: str) -> int | None:
    """빈 값은 None으로, 정수 문자열은 int로 바꿉니다."""
    정리한_값 = 값.strip()
    if 정리한_값 == "":
        return None
    try:
        # 현재 샘플에는 쉼표가 없으므로 쉼표가 들어간 숫자는 임의 처리하지 않습니다.
        if "," in 정리한_값:
            raise ValueError
        return int(정리한_값)
    except ValueError as 오류:
        raise 작업오류(
            f"{파일명} {행번호}행의 '{열이름}' 값이 정수가 아닙니다: {값!r}"
        ) from 오류


def 행_변환(행: list[str], 파일명: str, 행번호: int) -> list[Any]:
    """CSV 한 행을 엑셀에 기록할 자료형으로 바꿉니다."""
    결과: list[Any] = []
    for 열번호, 값 in enumerate(행):
        열이름 = 예상_제목[열번호]
        if 열번호 == 0:
            if 값.strip() == "":
                결과.append(None)
                continue
            try:
                결과.append(datetime.strptime(값.strip(), "%Y-%m-%d").date())
            except ValueError as 오류:
                raise 작업오류(
                    f"{파일명} {행번호}행의 '매출일' 형식이 YYYY-MM-DD가 아닙니다: {값!r}"
                ) from 오류
        elif 열번호 in 정수_열:
            결과.append(정수로_변환(값, 파일명, 행번호, 열이름))
        elif 열번호 in 문자_열:
            결과.append(값.strip() or None)
        else:
            결과.append(값.strip() or None)
    return 결과


def csv_전체_검사(csv_파일들: list[Path], 엑셀_제목: list[str]) -> list[list[Any]]:
    """모든 CSV를 먼저 검사하고 변환한 데이터를 메모리에 모읍니다."""
    전체_데이터: list[list[Any]] = []

    for 파일 in csv_파일들:
        try:
            with 파일.open("r", encoding="utf-8-sig", newline="") as 열린_파일:
                읽기 = csv.reader(열린_파일, strict=True)
                try:
                    제목 = next(읽기)
                except StopIteration as 오류:
                    raise 작업오류(f"빈 CSV 파일입니다: {파일.name}") from 오류

                if 제목 != 엑셀_제목:
                    raise 작업오류(
                        f"{파일.name}의 제목 또는 열 순서가 엑셀과 다릅니다.\n"
                        f"예상: {엑셀_제목}\n실제: {제목}"
                    )

                for 행번호, 행 in enumerate(읽기, start=2):
                    if len(행) != len(엑셀_제목):
                        raise 작업오류(
                            f"{파일.name} {행번호}행의 열 개수가 다릅니다. "
                            f"예상 {len(엑셀_제목)}개, 실제 {len(행)}개"
                        )
                    전체_데이터.append(행_변환(행, 파일.name, 행번호))
        except 작업오류:
            raise
        except UnicodeDecodeError as 오류:
            raise 작업오류(f"{파일.name}을 UTF-8 형식으로 읽을 수 없습니다: {오류}") from 오류
        except (OSError, csv.Error) as 오류:
            raise 작업오류(f"{파일.name}을 읽는 중 오류가 발생했습니다: {오류}") from 오류

    return 전체_데이터


def 완료_파일_충돌_검사(csv_파일들: list[Path], 완료_폴더: Path) -> None:
    """완료 폴더에 같은 이름이 하나라도 있는지 변경 전에 검사합니다."""
    충돌_파일 = [파일.name for 파일 in csv_파일들 if (완료_폴더 / 파일.name).exists()]
    if 충돌_파일:
        이름들 = ", ".join(충돌_파일)
        raise 작업오류(f"completed 폴더에 같은 이름의 파일이 있습니다: {이름들}")


def 셀_서식_복사(원본: Cell, 대상: Cell) -> None:
    """셀 값은 제외하고 기존 서식만 복사합니다."""
    if 원본.has_style:
        대상._style = copy(원본._style)
    if 원본.number_format:
        대상.number_format = 원본.number_format
    대상.font = copy(원본.font)
    대상.fill = copy(원본.fill)
    대상.border = copy(원본.border)
    대상.alignment = copy(원본.alignment)
    대상.protection = copy(원본.protection)


def 마지막_데이터_행(시트: Any) -> int:
    """서식이 아니라 실제 값이 들어 있는 마지막 행을 찾습니다."""
    for 행번호 in range(시트.max_row, 1, -1):
        if any(시트.cell(행번호, 열).value is not None for 열 in range(1, len(예상_제목) + 1)):
            return 행번호
    return 1


def 엑셀_갱신(엑셀_경로: Path, 데이터: list[list[Any]]) -> tuple[Any, int]:
    """통합 문서를 메모리에서 갱신하며 아직 파일에는 저장하지 않습니다."""
    try:
        통합문서 = load_workbook(엑셀_경로, data_only=False)
    except Exception as 오류:
        raise 작업오류(f"엑셀 파일을 다시 여는 중 오류가 발생했습니다: {오류}") from 오류

    try:
        시트 = 통합문서[시트_이름]
        기존_마지막_행 = 마지막_데이터_행(시트)

        for 행번호 in range(2, 기존_마지막_행 + 1):
            for 열번호 in range(1, len(예상_제목) + 1):
                시트.cell(행번호, 열번호).value = None

        for 데이터_번호, 행 in enumerate(데이터, start=2):
            # 기존 서식 범위보다 데이터가 길면 바로 위 행의 서식을 복사합니다.
            if 데이터_번호 > 시트.max_row:
                for 열번호 in range(1, len(예상_제목) + 1):
                    셀_서식_복사(
                        시트.cell(데이터_번호 - 1, 열번호),
                        시트.cell(데이터_번호, 열번호),
                    )
                이전_높이 = 시트.row_dimensions[데이터_번호 - 1].height
                if 이전_높이 is not None:
                    시트.row_dimensions[데이터_번호].height = 이전_높이

            for 열번호, 값 in enumerate(행, start=1):
                시트.cell(데이터_번호, 열번호).value = 값

        # 엑셀 프로그램에서 열었을 때 기존 수식이 다시 계산되도록 설정합니다.
        통합문서.calculation.fullCalcOnLoad = True
        통합문서.calculation.forceFullCalc = True
        return 통합문서, len(데이터)
    except Exception:
        통합문서.close()
        raise


def 임시_파일_검증(임시_경로: Path, 예상_데이터: list[list[Any]]) -> None:
    """저장한 임시 엑셀을 다시 열어 제목과 입력값을 확인합니다."""
    try:
        확인용 = load_workbook(임시_경로, read_only=True, data_only=False)
    except Exception as 오류:
        raise 작업오류(f"임시 엑셀 파일을 다시 열 수 없습니다: {오류}") from 오류

    try:
        if 시트_이름 not in 확인용.sheetnames:
            raise 작업오류("임시 엑셀 파일에서 '매출' 시트를 찾을 수 없습니다.")
        시트 = 확인용[시트_이름]
        실제_제목 = [시트.cell(1, 열).value for 열 in range(1, len(예상_제목) + 1)]
        if 실제_제목 != 예상_제목:
            raise 작업오류("임시 엑셀 파일의 제목이 변경되었습니다.")

        for 행번호, 예상_행 in enumerate(예상_데이터, start=2):
            실제_행 = [시트.cell(행번호, 열).value for 열 in range(1, len(예상_제목) + 1)]
            # openpyxl은 저장한 date를 다시 읽을 때 datetime으로 반환할 수 있습니다.
            if 실제_행 and isinstance(실제_행[0], datetime):
                실제_행[0] = 실제_행[0].date()
            if 실제_행 != 예상_행:
                raise 작업오류(f"임시 엑셀 파일의 {행번호}행이 원본 CSV 데이터와 다릅니다.")
    finally:
        확인용.close()


def 안전하게_저장(통합문서: Any, 엑셀_경로: Path, 데이터: list[list[Any]]) -> None:
    """임시 파일 저장과 검증이 끝난 뒤 원본 엑셀을 교체합니다."""
    임시_파일 = tempfile.NamedTemporaryFile(
        prefix="매출_임시_",
        suffix=".xlsx",
        dir=엑셀_경로.parent,
        delete=False,
    )
    임시_경로 = Path(임시_파일.name)
    임시_파일.close()

    try:
        통합문서.save(임시_경로)
        통합문서.close()
        임시_파일_검증(임시_경로, 데이터)
        os.replace(임시_경로, 엑셀_경로)
    except 작업오류:
        raise
    except Exception as 오류:
        raise 작업오류(f"엑셀 파일을 저장하는 중 오류가 발생했습니다: {오류}") from 오류
    finally:
        try:
            통합문서.close()
        except Exception:
            pass
        임시_경로.unlink(missing_ok=True)


def csv_이동(csv_파일들: list[Path], 완료_폴더: Path) -> int:
    """엑셀 저장이 끝난 뒤 CSV를 completed 폴더로 이동합니다."""
    완료_폴더.mkdir(exist_ok=True)
    이동한_파일: list[str] = []

    for 파일 in csv_파일들:
        목적지 = 완료_폴더 / 파일.name
        if 목적지.exists():
            raise 작업오류(
                "CSV 이동 중 같은 이름의 파일이 발견되었습니다. "
                f"이미 이동한 파일: {이동한_파일 or '없음'}, 실패 파일: {파일.name}"
            )
        try:
            shutil.move(str(파일), str(목적지))
            이동한_파일.append(파일.name)
        except Exception as 오류:
            raise 작업오류(
                f"CSV 이동에 실패했습니다. 이미 이동한 파일: {이동한_파일 or '없음'}, "
                f"실패 파일: {파일.name}, 원인: {오류}"
            ) from 오류
    return len(이동한_파일)


def 실행() -> None:
    """기획서 순서에 따라 전체 작업을 실행합니다."""
    엑셀_경로, csv_폴더, 완료_폴더 = 작업_경로_만들기()
    csv_파일들 = csv_목록_찾기(csv_폴더)

    # 이 구간이 끝날 때까지 어떠한 파일이나 폴더도 변경하지 않습니다.
    엑셀_제목 = 엑셀_제목_검사(엑셀_경로)
    완료_파일_충돌_검사(csv_파일들, 완료_폴더)
    데이터 = csv_전체_검사(csv_파일들, 엑셀_제목)

    통합문서, 데이터_행_수 = 엑셀_갱신(엑셀_경로, 데이터)
    안전하게_저장(통합문서, 엑셀_경로, 데이터)
    이동_파일_수 = csv_이동(csv_파일들, 완료_폴더)

    print("매출 데이터 통합을 완료했습니다.")
    print(f"- 읽은 CSV 파일: {len(csv_파일들)}개")
    print(f"- 입력한 데이터: {데이터_행_수}행")
    print(f"- 저장한 엑셀: {엑셀_경로}")
    print(f"- 이동한 CSV 파일: {이동_파일_수}개")


def main() -> int:
    try:
        실행()
        return 0
    except 작업오류 as 오류:
        print(f"작업을 중단했습니다.\n- 원인: {오류}", file=sys.stderr)
        return 1
    except Exception as 오류:
        print(f"예상하지 못한 오류로 작업을 중단했습니다.\n- 원인: {오류}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
