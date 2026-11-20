"""외부 API를 호출하지 않는 기본 단위 테스트입니다."""

import tempfile
import unittest
from pathlib import Path

from meeting_analyzer import _split_content
from notion_reader import NotionReaderError, extract_page_id
from output_writer import save_markdown


class ProgramTest(unittest.TestCase):
    def test_notion_url에서_page_id를_추출합니다(self) -> None:
        url = "https://app.notion.com/p/Meeting-41d3b7f5194a808d8c80c5265bc538c0?source=copy_link"
        self.assertEqual(extract_page_id(url), "41d3b7f5-194a-808d-8c80-c5265bc538c0")

    def test_notion이_아닌_주소를_거부합니다(self) -> None:
        with self.assertRaises(NotionReaderError):
            extract_page_id("https://example.com/41d3b7f5194a808d8c80c5265bc538c0")

    def test_긴_내용을_나눕니다(self) -> None:
        chunks = _split_content(("가상 회의 내용\n\n") * 5_000)
        self.assertGreater(len(chunks), 1)
        self.assertTrue(all(len(chunk) <= 24_100 for chunk in chunks))

    def test_기존_결과를_덮어쓰지_않습니다(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir)
            first = save_markdown("첫 번째", output_dir)
            second = save_markdown("두 번째", output_dir)
            self.assertNotEqual(first, second)
            self.assertEqual(first.read_text(encoding="utf-8"), "첫 번째\n")


if __name__ == "__main__":
    unittest.main()
