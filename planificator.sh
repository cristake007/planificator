#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${ROOT_DIR}/.venv"

usage() {
  cat <<'USAGE'
Usage: ./planificator.sh [--init | --install | --server]

Commands:
  --init      Create a local virtual environment at .venv
  --install   Upgrade pip and install dependencies from requirements.txt
  --server    Run the Flask app on 0.0.0.0:5000
USAGE
}

ensure_venv_exists() {
  if [[ ! -d "${VENV_DIR}" ]]; then
    echo "Virtual environment not found at ${VENV_DIR}. Run --init first."
    exit 1
  fi
}

activate_venv() {
  # shellcheck disable=SC1091
  source "${VENV_DIR}/bin/activate"
}

case "${1:-}" in
  --init)
    if [[ -d "${VENV_DIR}" ]]; then
      echo "Virtual environment already exists at ${VENV_DIR}."
    else
      python3 -m venv "${VENV_DIR}"
      echo "Virtual environment created at ${VENV_DIR}."
    fi
    echo "To activate manually, run: source .venv/bin/activate"
    ;;
  --install)
    ensure_venv_exists
    activate_venv
    python -m pip install --upgrade pip
    pip install -r "${ROOT_DIR}/requirements.txt"
    ;;
  --server)
    ensure_venv_exists
    activate_venv
    exec python "${ROOT_DIR}/run.py"
    ;;
  *)
    usage
    exit 1
    ;;
esac
