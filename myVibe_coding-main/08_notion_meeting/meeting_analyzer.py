"""OpenAI Responses API를 사용해 회의록을 구조적으로 분석합니다."""

from openai import OpenAI, OpenAIError


MAX_CHUNK_CHARS = 24_000


class MeetingAnalyzerError(Exception):
    """사용자에게 안내할 수 있는 회의록 분석 오류입니다."""


def _split_content(content: str) -> list[str]:
    if len(content) <= MAX_CHUNK_CHARS:
        return [content]

    chunks: list[str] = []
    current: list[str] = []
    current_length = 0
    for paragraph in content.split("\n\n"):
        addition = len(paragraph) + 2
        if current and current_length + addition > MAX_CHUNK_CHARS:
            chunks.append("\n\n".join(current))
            current = []
            current_length = 0
        current.append(paragraph)
        current_length += addition
    if current:
        chunks.append("\n\n".join(current))
    return chunks


def _call_model(client: OpenAI, model: str, prompt: str) -> str:
    try:
        response = client.responses.create(model=model, input=prompt)
    except OpenAIError as exc:
        raise MeetingAnalyzerError("AI 분석 요청이 실패했습니다. API 키, 모델명, 사용 한도를 확인하세요.") from exc
    text = response.output_text.strip()
    if not text:
        raise MeetingAnalyzerError("AI가 분석 결과를 반환하지 않았습니다.")
    return text


def analyze_meeting(
    title: str,
    source_url: str,
    content: str,
    api_key: str,
    model: str,
    warnings: list[str] | None = None,
) -> str:
    client = OpenAI(api_key=api_key)
    chunks = _split_content(content)

    if len(chunks) == 1:
        analysis_source = content
    else:
        partials: list[str] = []
        for index, chunk in enumerate(chunks, start=1):
            partials.append(
                _call_model(
                    client,
                    model,
                    f"""다음은 긴 회의록의 {index}/{len(chunks)} 부분입니다.
원문에 있는 사실만 사용하여 수치, 날짜, 결정사항, 미결 사항, 담당자, 할 일, 기한, 체크박스 상태를 빠짐없이 추출하세요.
추정하지 말고 불명확한 값은 '확인 필요'라고 쓰세요.

{chunk}""",
                )
            )
        analysis_source = "\n\n".join(partials)

    warning_text = "\n".join(f"- {item}" for item in (warnings or [])) or "- 없음"
    prompt = f"""당신은 회의록을 운영 업무에 사용할 수 있게 정리하는 분석가입니다.
아래 자료에 없는 사실은 절대 만들지 마세요. 제안, 검토안, 확정된 결정사항을 구분하세요.
담당자·기한·상태가 없으면 추정하지 말고 반드시 '확인 필요'라고 표시하세요.
원문의 체크박스가 `[ ]`이면 상태를 `미완료`, `[x]`이면 `완료`로 표시하세요.
체크박스 자체가 없고 완료 여부를 알 수 없을 때만 상태를 `확인 필요`로 표시하세요.
같은 업무가 충돌하면 임의로 합치지 말고 확인 필요 항목에 기록하세요.
숫자와 날짜는 원문을 그대로 보존하세요. 출력은 한국어 Markdown만 사용하세요.

반드시 다음 순서와 제목을 사용하세요.
# 회의록 분석 결과
## 기본 정보
- 회의록 제목
- 원본 페이지
## 회의 요약
## 현안 및 당면 문제
문제, 영향, 긴급도, 현재 대응, 상태 열이 있는 표로 작성하세요.
## 결정사항
## 미결 사항
## 담당자별 액션 아이템
담당자, 소속/역할, 해야 할 일, 기한, 상태, 근거 열이 있는 표로 작성하세요.
## 확인 필요 항목
## 분석 시 주의사항

회의록 제목: {title}
원본 페이지: {source_url}
읽기 경고:
{warning_text}

분석할 내용:
{analysis_source}"""
    return _call_model(client, model, prompt)
