"""Notion 회의록을 읽고 구조화된 분석 결과를 만드는 실행 파일입니다."""

import os
import sys
from pathlib import Path

from dotenv import load_dotenv

from meeting_analyzer import MeetingAnalyzerError, analyze_meeting
from notion_reader import NotionReaderError, read_notion_page
from output_writer import save_markdown


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent


def _require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f".env에 {name} 값이 없습니다.")
    return value


def main() -> int:
    # 이 실습 프로젝트는 공통 비밀정보를 프로젝트 루트의 .env에서 읽습니다.
    load_dotenv(PROJECT_DIR / ".env")

    print("Notion 회의록 요약 및 액션 아이템 추출 프로그램")
    page_url = input(
        "분석할 Notion 회의록 페이지 주소를 입력하세요.\n"
        "비워 두면 .env의 NOTION_PAGE_ID를 사용합니다.\n> "
    ).strip()
    if not page_url:
        page_url = os.getenv("NOTION_PAGE_ID", "").strip()

    if not page_url:
        print("오류: 페이지 주소를 입력하거나 .env에 NOTION_PAGE_ID를 설정해야 합니다.")
        return 1

    try:
        notion_token = _require_env("NOTION_TOKEN")
        openai_api_key = _require_env("OPENAI_API_KEY")

        print("회의록을 읽고 있습니다...")
        page = read_notion_page(page_url, notion_token)

        print("회의 내용을 분석하고 있습니다...")
        result = analyze_meeting(
            title=page.title,
            source_url=page_url,
            content=page.content,
            api_key=openai_api_key,
            model=(
                os.getenv("OPENAI_DEFAULT_MODEL")
                or os.getenv("OPENAI_MODEL")
                or "gpt-6.2-astra"
            ).strip(),
            warnings=page.warnings,
        )

        output_path = save_markdown(result, BASE_DIR / "output")
        print("\n" + result)
        print(f"\n분석이 완료되었습니다.\n저장 위치: {output_path}")
        return 0
    except (ValueError, NotionReaderError, MeetingAnalyzerError) as exc:
        print(f"오류: {exc}")
        return 1
    except KeyboardInterrupt:
        print("\n사용자가 실행을 중단했습니다.")
        return 130
    except Exception:
        print("오류: 예상하지 못한 문제가 발생했습니다. 비밀정보 보호를 위해 상세 내용은 출력하지 않습니다.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
