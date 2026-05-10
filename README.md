# 📊 AI 벤치마킹 보고서 자동 작성 시스템

> HuggingFace LLM (`google/gemma-4-26B-A4B-it`) × Multi-source VOC Collector × Streamlit × MS Office Export

---

## 🆕 주요 기능

- 멀티소스 VOC 자동 수집: DuckDuckGo 웹/뉴스, Google News RSS
- VOC 중복 제거, 도메인 추출, 채널 분류, 카테고리/감성 힌트 자동 태깅
- LLM 기반 VOC 심층 분석 및 제품 사양 비교
- 보고서 다운로드: Markdown, HTML, **Word(.docx)**, **PowerPoint(.pptx)**, JSON

## 🗂️ 프로젝트 구조

```
benchmarking_app/
├── app.py                  ← 메인 Streamlit 애플리케이션
├── requirements.txt        ← Python 패키지 의존성
├── .env.example            ← 환경변수 템플릿
├── .env                    ← 실제 환경변수 (Git 제외)
├── Dockerfile              ← Docker 이미지 빌드
├── docker-compose.yml      ← 컨테이너 오케스트레이션
├── run.sh                  ← 로컬/Docker 통합 실행 스크립트
├── .gitignore
└── .streamlit/
    └── config.toml         ← Streamlit 테마 & 서버 설정
```

---

## ⚙️ 사전 준비

### 1. HuggingFace API 토큰 발급

1. [huggingface.co](https://huggingface.co) 회원가입
2. `Settings → Access Tokens → New Token` (Read 권한)
3. 아래 모델 페이지에서 라이선스 동의 **(필수)**:
   → [google/gemma-4-26B-A4B-it](https://huggingface.co/google/gemma-4-26B-A4B-it)

### 2. 환경변수 설정

```bash
cp .env.example .env
# .env 파일을 열어 HF_TOKEN 값 입력
```

---

## 🚀 실행 방법

### A. 로컬 직접 실행 (개발/테스트)

```bash
chmod +x run.sh
./run.sh
# → http://localhost:8501
```

수동 설치:
```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -U -r requirements.txt
streamlit run app.py
```

### B. Docker (서버 배포)

```bash
./run.sh docker        # 빌드 & 백그라운드 실행
./run.sh stop          # 중지
docker compose logs -f # 로그 확인
```

### C. Streamlit Community Cloud (무료 배포)

1. GitHub 레포에 Push
2. [share.streamlit.io](https://share.streamlit.io) → `New app` → 레포 선택
3. `Advanced settings → Secrets`:
   ```toml
   HF_TOKEN = "hf_실제토큰값"
   ```
4. `Deploy!`

---

## 📋 기능 설명 (4단계 워크플로우)

| 단계 | 탭 | 기능 | 소요 시간 |
|------|-----|------|-----------|
| Step 1 | 가이드 추출 | 기존 보고서 PDF/DOCX 업로드 → LLM이 공통 구조·항목 추출 | 1~3분 |
| Step 2 | VOC 수집 | 자사·경쟁사 입력 → 웹/뉴스/RSS 멀티소스 수집 → 중복 제거 → 카테고리·감성 분류 → LLM 심층 분석 | 2~7분 |
| Step 3 | 사양 비교 | 사양 직접 입력 또는 웹 수집 → AI 비교 매트릭스·갭 분석 | 1~2분 |
| Step 4 | 보고서 생성 | 1~3단계 데이터 통합 → McKinsey급 보고서 자동 작성 | 2~5분 |


## 🆕 VOC 자동 수집 업그레이드 내용

이번 버전의 Step 2 VOC 수집부는 단순 검색 결과 나열이 아니라, 보고서 작성에 바로 쓸 수 있는 정제 VOC 데이터셋을 생성하도록 개선되었습니다.

- 수집 소스 선택: 일반 웹, DuckDuckGo 뉴스, Google News RSS
- 수집 깊이 선택: 기본 / 표준 / 심층
- 검색 기간 선택: 전체, 최근 1일, 1주, 1개월, 1년
- 중복 제거: URL 추적 파라미터 제거, 제목 기반 중복 제거
- 자동 분류: 성능/속도, 배터리/발열, 카메라/화질, 가격/가성비, 소프트웨어/UI, A/S 등
- 감성 힌트: 긍정, 부정, 혼합, 개선요청, 중립
- 신뢰도 보강: 출처 도메인, 채널, relevance score 부여
- 분석 결과 시각화: 카테고리 분포와 감성 분포 차트
- CSV 다운로드: 수집된 정제 VOC 테이블을 `utf-8-sig` CSV로 저장 가능
- 본문 일부 보강 옵션: 검색 snippet이 짧을 때 URL 본문 일부를 추가 수집

본문 일부 보강 옵션은 더 풍부한 분석에 유리하지만, 일부 사이트가 크롤링을 막거나 응답이 느릴 수 있으므로 Streamlit Cloud에서는 기본값 OFF를 권장합니다.

---

## 🔧 모델 변경

사이드바 또는 `.env`에서:

```env
# 대안 모델 예시
HF_MODEL_ID=mistralai/Mixtral-8x7B-Instruct-v0.1
HF_MODEL_ID=meta-llama/Llama-3.1-70B-Instruct
HF_MODEL_ID=Qwen/Qwen2.5-72B-Instruct
```

---

## ❓ 트러블슈팅

| 증상 | 원인 | 해결 |
|------|------|------|
| `401 Unauthorized` | 토큰 오류 | HF 토큰 재확인 |
| `Model not found` | 접근 권한 없음 | 모델 페이지에서 라이선스 동의 |
| `Too many requests` | API 한도 초과 | 잠시 후 재시도 |
| VOC 수집 0건 | DuckDuckGo 차단 | VPN 변경 또는 검색어 수정 |
| PDF 텍스트 없음 | 스캔 이미지 PDF | OCR 전처리 후 재업로드 |

---

## 🏗️ 아키텍처

```
사용자 브라우저 (Streamlit UI)
        │
  ┌─────┴──────────────────────────┐
  │         app.py (메인 엔진)      │
  │  ┌──────────┐ ┌─────────────┐  │
  │  │File Parser│ │Web Crawler  │  │
  │  │PDF·DOCX  │ │DuckDuckGo  │  │
  │  └──────┬───┘ └──────┬──────┘  │
  │         └──────┬──────┘         │
  │    ┌───────────▼──────────┐     │
  │    │  LLM Orchestrator    │     │
  │    │  Prompt Engineering  │     │
  │    └───────────┬──────────┘     │
  └────────────────┼────────────────┘
                   ▼
      HuggingFace Inference API
      google/gemma-4-26B-A4B-it
                   │
          ┌────────┴────────┐
    JSON Guide  VOC분석  사양비교  보고서MD
```

---

*AI 벤치마킹 보고서 시스템 v2.0 | Streamlit + HuggingFace LLM*


## 긴급 수정: Gemma/ZAYA 모델의 `text-generation` 오류

다음 오류가 표시되는 경우:

```text
Model google/gemma-4-26B-A4B-it is not supported for task text-generation and provider novita. Supported task: conversational.
```

원인은 모델/provider가 `text_generation()`이 아니라 `chat_completion()` / `chat.completions.create()` 방식만 지원하기 때문입니다.
이 버전에서는 `app.py`의 LLM 호출부를 conversational 방식으로 수정했고, 사이드바에서 Inference Provider를 `auto`, `Novita AI`, `HF Inference` 등으로 바꿔 연결 테스트할 수 있게 했습니다.

권장 재설치:

```bash
pip install -U -r requirements.txt
streamlit run app.py
```

Streamlit Cloud에서는 Secrets에 아래처럼 넣을 수 있습니다.

```toml
HF_TOKEN = "hf_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
HF_MODEL_ID = "google/gemma-4-26B-A4B-it"
HF_PROVIDER = "auto"
```


### DOCX/PPTX 다운로드

보고서 생성 후 Step 4 하단에서 다음 형식으로 바로 다운로드할 수 있습니다.

- Word DOCX: 본문형 보고서, 표 포함
- PowerPoint PPTX: 발표용 요약 덱, VOC/사양 비교 슬라이드 포함

