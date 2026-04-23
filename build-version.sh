#!/bin/sh
set -eu

escape_js() {
  printf '%s' "$1" | sed 's/\\/\\\\/g; s/"/\\"/g'
}

branch="${BRANCH:-${COMMIT_REF:-}}"
if [ -z "${branch}" ]; then
  branch="$(git branch --show-current 2>/dev/null || true)"
fi

hash="$(git rev-parse --short HEAD 2>/dev/null || true)"
if [ -z "${hash}" ]; then
  hash="${COMMIT_REF:-}"
  if [ -n "${hash}" ]; then
    hash="$(printf '%s' "$hash" | cut -c1-7)"
  fi
fi

if [ -z "${hash}" ]; then
  hash="unknown"
fi

branch_js="null"
if [ -n "${branch}" ]; then
  branch_js="\"$(escape_js "$branch")\""
fi

hash_js="\"$(escape_js "$hash")\""

cat > version.js <<EOF
window.__PAGE_BAKER_VERSION__ = {
  branch: ${branch_js},
  hash: ${hash_js}
};
EOF

if [ -f index.html ]; then
  sed -i "s#<script src=\"\\./version\\.js\"></script>#<script src=\"./version.js?v=${hash}\"></script>#" index.html
fi
