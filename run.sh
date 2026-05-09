#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
# AI 벤치마킹 보고서 시스템  ·  실행 스크립트
# 사용법:
#   ./run.sh            — 로컬 venv로 직접 실행
#   ./run.sh docker     — Docker 컨테이너로 빌드 & 실행
#   ./run.sh install    — 의존성만 설치
#   ./run.sh stop       — Docker 컨테이너 중지
# ─────────────────────────────────────────────────────────────────────────────
set -euo pipefail

VENV_DIR=".venv"
PORT=8501

RED='\033[0;31m'; GRN='\033[0;32m'; YLW='\033[1;33m'
BLU='\033[0;34m'; PRP='\033[0;35m'; CYN='\033[0;36m'; NC='\033[0m'

banner() {
  echo -e "${PRP}"
  echo "  ╔══════════════════════════════════════════════════════════╗"
  echo "  ║      📊  AI 벤치마킹 보고서 자동 작성 시스템           ║"
  echo "  ║      HuggingFace LLM · DuckDuckGo · Streamlit           ║"
  echo "  ╚══════════════════════════════════════════════════════════╝"
  echo -e "${NC}"
}

log()  { echo -e "${GRN}[✓]${NC} $*"; }
warn() { echo -e "${YLW}[!]${NC} $*"; }
err()  { echo -e "${RED}[✗]${NC} $*" >&2; exit 1; }
info() { echo -e "${CYN}[i]${NC} $*"; }

check_env() {
  if [[ ! -f ".env" ]]; then
    warn ".env 파일이 없습니다. .env.example 에서 복사합니다…"
    cp .env.example .env
    warn "⚠️  .env 파일을 열어 HF_TOKEN 값을 반드시 입력하세요!"
  fi
  set -a
  # shellcheck disable=SC1091
  source .env
  set +a
  if [[ -z "${HF_TOKEN:-}" || "${HF_TOKEN}" == hf_xxx* ]]; then
    warn "HF_TOKEN이 설정되지 않았습니다. .env 파일을 편집 후 재실행하세요."
  else
    log "HF_TOKEN 확인됨 (${#HF_TOKEN}자)"
  fi
}

install_deps() {
  info "Python 패키지 설치/업그레이드 중…"
  pip install --upgrade pip --quiet
  pip install -r requirements.txt --quiet
  log "패키지 설치 완료"
}

run_local() {
  banner; check_env
  PYTHON=$(command -v python3 || command -v python || err "Python3가 없습니다.")
  info "Python: $("$PYTHON" --version)"
  if [[ ! -d "${VENV_DIR}" ]]; then
    info "가상환경 생성 중…"
    "$PYTHON" -m venv "${VENV_DIR}"
    log "가상환경 생성 완료"
  fi
  # shellcheck disable=SC1091
  source "${VENV_DIR}/bin/activate"
  log "가상환경 활성화"
  install_deps
  mkdir -p .streamlit
  info "Streamlit 앱 시작 중…"
  echo -e "\n  ${GRN}🌐 접속:${NC} ${BLU}http://localhost:${PORT}${NC}\n"
  streamlit run app.py \
    --server.port="${PORT}" \
    --server.address="0.0.0.0" \
    --server.headless=true \
    --browser.gatherUsageStats=false
}

run_docker() {
  banner; check_env
  command -v docker &>/dev/null || err "Docker가 설치되어 있지 않습니다."
  COMPOSE="docker compose"
  command -v docker compose &>/dev/null || COMPOSE="docker-compose"
  info "Docker 이미지 빌드 중… (첫 실행 시 수 분 소요)"
  $COMPOSE build --no-cache
  info "컨테이너 시작 중…"
  $COMPOSE up -d
  echo -e "\n  ${GRN}🌐 접속:${NC} ${BLU}http://localhost:${PORT}${NC}\n"
  log "로그: docker compose logs -f benchmarking-app"
  log "중지: ./run.sh stop"
}

stop_docker() {
  COMPOSE="docker compose"
  command -v docker compose &>/dev/null || COMPOSE="docker-compose"
  info "컨테이너 중지 중…"
  $COMPOSE down
  log "컨테이너 중지 완료"
}

case "${1:-local}" in
  local|"") run_local   ;;
  docker)   run_docker  ;;
  install)  install_deps ;;
  stop)     stop_docker ;;
  *)
    echo "사용법: $0 [local|docker|install|stop]"
    exit 1 ;;
esac
