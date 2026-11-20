"""Notion API에서 페이지와 하위 블록을 읽어 텍스트로 변환합니다."""

from dataclasses import dataclass, field
import re
from typing import Any
from urllib.parse import urlparse

import requests


NOTION_API_BASE = "https://api.notion.com/v1"
NOTION_VERSION = "2026-03-11"
MAX_DEPTH = 20


class NotionReaderError(Exception):
    """사용자에게 안내할 수 있는 Notion 읽기 오류입니다."""


@dataclass
class NotionPage:
    title: str
    content: str
    warnings: list[str] = field(default_factory=list)


def extract_page_id(value: str) -> str:
    """Notion 주소 또는 페이지 ID에서 32자리 페이지 ID를 추출합니다."""
    value = value.strip()
    parsed = urlparse(value)
    if parsed.scheme and parsed.netloc and "notion" not in parsed.netloc.lower():
        raise NotionReaderError("Notion 페이지 주소만 사용할 수 있습니다.")

    compact = value.replace("-", "")
    matches = re.findall(r"(?<![0-9a-fA-F])[0-9a-fA-F]{32}(?![0-9a-fA-F])", compact)
    if not matches:
        raise NotionReaderError("페이지 ID를 찾을 수 없습니다. Notion에서 '링크 복사'로 얻은 주소를 입력하세요.")

    raw = matches[-1].lower()
    return f"{raw[:8]}-{raw[8:12]}-{raw[12:16]}-{raw[16:20]}-{raw[20:]}"


def _rich_text(items: list[dict[str, Any]]) -> str:
    return "".join(str(item.get("plain_text", "")) for item in items)


def _request_json(url: str, token: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
    try:
        response = requests.get(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Notion-Version": NOTION_VERSION,
                "Content-Type": "application/json",
            },
            params=params,
            timeout=30,
        )
    except requests.RequestException as exc:
        raise NotionReaderError("Notion 서버에 연결하지 못했습니다. 인터넷 연결을 확인하세요.") from exc

    if response.status_code == 401:
        raise NotionReaderError("Notion 토큰이 올바르지 않습니다.")
    if response.status_code == 403:
        raise NotionReaderError("이 페이지를 읽을 권한이 없습니다. Notion에서 연결에 페이지 접근 권한을 부여하세요.")
    if response.status_code == 404:
        raise NotionReaderError("페이지를 찾지 못했습니다. 주소와 연결 권한을 확인하세요.")
    if not response.ok:
        raise NotionReaderError(f"Notion API 요청이 실패했습니다. 상태 코드: {response.status_code}")

    try:
        return response.json()
    except ValueError as exc:
        raise NotionReaderError("Notion 응답을 읽을 수 없습니다.") from exc


def _page_title(page: dict[str, Any]) -> str:
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            title = _rich_text(prop.get("title", []))
            if title:
                return title
    return "제목 없는 회의록"


def _block_line(block: dict[str, Any], warnings: list[str]) -> str:
    block_type = block.get("type", "unknown")
    data = block.get(block_type, {})
    text = _rich_text(data.get("rich_text", []))

    if block_type == "heading_1":
        return f"# {text}"
    if block_type == "heading_2":
        return f"## {text}"
    if block_type == "heading_3":
        return f"### {text}"
    if block_type == "bulleted_list_item":
        return f"- {text}"
    if block_type == "numbered_list_item":
        return f"1. {text}"
    if block_type == "to_do":
        mark = "x" if data.get("checked") else " "
        return f"- [{mark}] {text}"
    if block_type == "quote":
        return f"> {text}"
    if block_type == "callout":
        return f"> {text}"
    if block_type == "code":
        language = data.get("language", "")
        return f"```{language}\n{text}\n```"
    if block_type == "divider":
        return "---"
    if block_type == "child_page":
        return f"[하위 페이지] {data.get('title', '')}"
    if block_type == "table_row":
        cells = [_rich_text(cell) for cell in data.get("cells", [])]
        return "| " + " | ".join(cells) + " |"
    if block_type in {"paragraph", "toggle"}:
        return text
    if block_type in {"image", "video", "file", "pdf", "audio", "bookmark", "embed"}:
        warnings.append(f"{block_type} 블록의 첨부 내용은 분석에서 제외되었습니다.")
        return f"[{block_type} 첨부 블록]"
    if block_type not in {"column", "column_list", "table", "synced_block"}:
        warnings.append(f"지원하지 않는 블록 형식이 있어 일부 내용이 누락될 수 있습니다: {block_type}")
    return text


def _read_children(block_id: str, token: str, warnings: list[str], depth: int = 0) -> list[str]:
    if depth > MAX_DEPTH:
        warnings.append("하위 블록 깊이가 너무 깊어 일부 내용을 생략했습니다.")
        return []

    lines: list[str] = []
    cursor: str | None = None
    while True:
        params: dict[str, Any] = {"page_size": 100}
        if cursor:
            params["start_cursor"] = cursor
        data = _request_json(f"{NOTION_API_BASE}/blocks/{block_id}/children", token, params)

        for block in data.get("results", []):
            line = _block_line(block, warnings)
            if line.strip():
                lines.append(line)
            if block.get("has_children"):
                lines.extend(_read_children(block["id"], token, warnings, depth + 1))

        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
        if not cursor:
            warnings.append("Notion 페이지의 다음 내용을 불러오지 못했습니다.")
            break
    return lines


def read_notion_page(page_url: str, token: str) -> NotionPage:
    page_id = extract_page_id(page_url)
    page = _request_json(f"{NOTION_API_BASE}/pages/{page_id}", token)
    warnings: list[str] = []
    lines = _read_children(page_id, token, warnings)
    content = "\n".join(lines).strip()
    if not content:
        raise NotionReaderError("분석할 회의록 내용이 없습니다.")
    return NotionPage(title=_page_title(page), content=content, warnings=sorted(set(warnings)))
