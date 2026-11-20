"""분석 결과를 기존 파일과 겹치지 않게 저장합니다."""

from datetime import datetime
from pathlib import Path


def save_markdown(content: str, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    base = output_dir / "회의록_분석.md"
    if not base.exists():
        target = base
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = output_dir / f"회의록_분석_{timestamp}.md"
        number = 2
        while target.exists():
            target = output_dir / f"회의록_분석_{timestamp}_{number}.md"
            number += 1
    target.write_text(content.rstrip() + "\n", encoding="utf-8")
    return target
