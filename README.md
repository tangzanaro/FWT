# 외래어 한글 표기 추천 시스템 (FWT)

`Foreign word transcription_Usage List.xlsx`와 `Foreign word transcription.rag.full.md`를 기반으로  
외래어 한글 표기를 추천하는 Streamlit 앱입니다.

## 주요 기능

- **용례집 정확 일치 우선 검색**
  - 입력 원어 + 언어권 기준으로 용례집에서 정확 일치 항목을 우선 반환합니다.
- **신규 단어 RAG + LLM 추천**
  - 용례집에 없는 단어는 RAG 규칙 문맥을 구성한 뒤 로컬 Ollama 모델(`gpt-oss:20b`)로 표기를 생성합니다.
- **근거/출처 표시**
  - 결과와 함께 `B#####/P#####/T###` 기반 참조 조항을 보여줍니다.
- **안전한 실패 처리**
  - Ollama 연결 실패, 출력 형식 오류, 규칙 부족 시 임의 표기를 강제하지 않고 실패 상태를 명확히 표시합니다.

## 프로젝트 구조

- `fwt.py` : 메인 Streamlit 앱 및 추천 파이프라인
- `test_fwt.py` : 파서/판정/실패 처리 테스트
- `Foreign word transcription_Usage List.xlsx` : 용례집 데이터
- `Foreign word transcription.rag.full.md` : 규칙 RAG 문서

## 실행 환경

- Python 3.10+
- Streamlit
- Pandas / Openpyxl
- Ollama
- 로컬 모델: `gpt-oss:20b`

## 설치

```bash
pip install streamlit pandas openpyxl pytest
```

Ollama 설치 후 모델 준비:

```bash
ollama pull gpt-oss:20b
```

## 실행 방법

1. Ollama 서버 실행

```bash
ollama serve
```

2. 앱 실행

```bash
streamlit run fwt.py
```

## 새로운 PC 온보딩 체크리스트

세팅이 이미 되어 있는 PC라면 아래만 확인하면 바로 실행할 수 있습니다.

1. 프로젝트 필수 파일 확인
   - `fwt.py`
   - `Foreign word transcription_Usage List.xlsx`
   - `Foreign word transcription.rag.full.md`
2. Python/패키지 확인
   - `python --version`
   - `python -c "import streamlit,pandas,openpyxl; print('ok')"`
3. Ollama/모델 확인
   - `ollama list`에서 `gpt-oss:20b` 존재 확인
4. 서버/앱 실행
   - `ollama serve`
   - `streamlit run fwt.py`

문제 발생 시 `## 문제 해결` 섹션을 먼저 확인해 주세요.

## 입력값

- **외래어 문자열**: 필수
- **언어권(국명/언어명)**: 필수
- **발음/IPA**: 선택
  - 입력하지 않아도 동작합니다.
  - 필요 시 모델이 발음을 추정하고, 불확실하면 상태에 표시합니다.

## 결과 상태 예시

- `용례집 정확 일치`
- `RAG 기반 신규 표기`
- `RAG 기반 신규 표기(발음 추정)`
- `발음 확인 필요`
- `언어권 확인 필요`
- `RAG 검색 실패`
- `Ollama 호출 실패`
- `표기 생성 실패`
- `표기 확정 실패`

## 테스트

```bash
pytest -q
```

## 문제 해결

### 1) `Connection refused` / Ollama 연결 실패

- Ollama 서버가 꺼져 있을 가능성이 큽니다.

```bash
ollama serve
```

- 모델 설치 여부 확인:

```bash
ollama list
```

### 2) 모델 출력 형식(JSON) 오류

- 앱에서 재시도/복구 파싱을 수행합니다.
- 계속 반복되면 Ollama와 모델 상태를 확인하고 앱을 재시작해 주세요.

## 참고

- 규칙 근거는 `Foreign word transcription.rag.full.md`를 기준으로 파싱합니다.
- `Lxxxx` ID가 아니라 실제 문서의 `B#####`, `P#####`, `T###` 식별자를 사용합니다.
