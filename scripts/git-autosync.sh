#!/bin/zsh

set -u

# Scheduled runs must fail clearly instead of waiting for an interactive password.
export GIT_TERMINAL_PROMPT=0

PROJECT_DIR="${1:-$PWD}"
LOG_DIR="$HOME/.codex/logs"
LOG_FILE="$LOG_DIR/fund-investment-autosync.log"
RUN_DIR="$HOME/.codex/run"
LOCK_DIR="$RUN_DIR/fund-investment-autosync.lock"

mkdir -p "$LOG_DIR" "$RUN_DIR"

log() {
  print -r -- "$(date '+%Y-%m-%d %H:%M:%S') $*" >> "$LOG_FILE"
}

if ! mkdir "$LOCK_DIR" 2>/dev/null; then
  if [[ -f "$LOCK_DIR/pid" ]] && kill -0 "$(<"$LOCK_DIR/pid")" 2>/dev/null; then
    log "Another sync is running; skipping."
    exit 0
  fi
  rm -f "$LOCK_DIR/pid"
  rmdir "$LOCK_DIR" 2>/dev/null || {
    log "Could not clear a stale sync lock; skipping."
    exit 0
  }
  mkdir "$LOCK_DIR" 2>/dev/null || {
    log "Could not acquire sync lock; skipping."
    exit 0
  }
fi
print -r -- "$$" > "$LOCK_DIR/pid"
trap 'rm -f "$LOCK_DIR/pid"; rmdir "$LOCK_DIR" 2>/dev/null' EXIT INT TERM

# Command-line tools do not always inherit macOS System Settings proxy values.
proxy_address=$(/usr/sbin/scutil --proxy | /usr/bin/awk '
  /HTTPSEnable : 1/ { enabled = 1 }
  /HTTPSProxy :/ { host = $3 }
  /HTTPSPort :/ { port = $3 }
  END { if (enabled && host && port) print host ":" port }
')
if [[ -n "$proxy_address" ]]; then
  export HTTP_PROXY="http://$proxy_address"
  export HTTPS_PROXY="http://$proxy_address"
  export http_proxy="$HTTP_PROXY"
  export https_proxy="$HTTPS_PROXY"
fi

cd "$PROJECT_DIR" 2>/dev/null || {
  log "Project directory is unavailable: $PROJECT_DIR"
  exit 1
}

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  log "Not a git repository: $PROJECT_DIR"
  exit 1
fi

if ! git remote get-url origin >/dev/null 2>&1; then
  log "No origin remote configured; skipping sync."
  exit 1
fi

branch=$(git branch --show-current)
if [[ -z "$branch" ]]; then
  log "No current branch; skipping sync."
  exit 1
fi

log "Starting sync for $branch."

if [[ -n "$(git status --porcelain)" ]]; then
  if ! git add -A >> "$LOG_FILE" 2>&1; then
    log "git add failed; sync stopped."
    exit 1
  fi
  if [[ -n "$(git diff --cached --name-only)" ]]; then
    if ! git commit -m "autosync: $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG_FILE" 2>&1; then
      log "git commit failed; sync stopped."
      exit 1
    fi
    log "Committed local changes."
  fi
fi

if ! git fetch origin "$branch" >> "$LOG_FILE" 2>&1; then
  log "Fetch failed; sync stopped without pushing."
  exit 1
fi

if git rev-parse --verify "origin/$branch" >/dev/null 2>&1; then
  if ! git pull --rebase --autostash origin "$branch" >> "$LOG_FILE" 2>&1; then
    log "Pull/rebase failed; sync stopped without pushing."
    if [[ -d "$(git rev-parse --git-path rebase-merge)" || -d "$(git rev-parse --git-path rebase-apply)" ]]; then
      if git rebase --abort >> "$LOG_FILE" 2>&1; then
        log "Aborted the incomplete rebase; local commits were preserved."
      else
        log "Could not abort the incomplete rebase; manual intervention is required."
      fi
    fi
    exit 1
  fi
else
  log "Remote branch origin/$branch does not exist; preparing its first push."
fi

if git push -u origin "$branch" >> "$LOG_FILE" 2>&1; then
  log "Pushed $branch successfully."
else
  log "Push failed; no force push was attempted."
  exit 1
fi
