# 회의결과보고서 자동생성기

회의 메모(대화체 원문)를 입력하면 Gemini가 개조식 단문형 회의결과보고서로 정리하고,  
Word 템플릿의 `{{ }}` 플레이스홀더를 자동으로 채워 `.docx` 파일로 내려받는 Streamlit 앱입니다.

## 구성 파일

- `app.py`: Streamlit 웹앱 본체
- `requirements.txt`: 의존성 목록
- `.gitignore`: Python/Streamlit용 무시 규칙

## 템플릿 플레이스홀더

현재 앱은 아래 키를 채웁니다.

- `{{title}}`
- `{{date}}`
- `{{location}}`
- `{{reporter}}`
- `{{attendees}}`
- `{{agenda}}`
- `{{content}}`
- `{{tasks}}`
- `{{references}}`

## 실행 방법

1) 가상환경 생성/활성화 (선택)
2) 패키지 설치

```bash
pip install -r requirements.txt
```

3) Streamlit secrets 설정  
`.streamlit/secrets.toml` 파일을 만들고 아래처럼 입력:

```toml
GEMINI_API_KEY = "여기에_본인_제미나이_API_키"
```

4) 앱 실행

```bash
streamlit run app.py
```

## 사용 순서

1. 회의 메모 원문 입력
2. `AI로 회의결과 정리` 버튼 클릭
3. 주요안건/결정사항 확인
4. AI 결과 직접 수정
5. `Word 회의결과보고서 생성` 후 다운로드
