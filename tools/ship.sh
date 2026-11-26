#!/usr/bin/env bash
# Ship a change: bump the version implied by the work since the last version
# change, commit it, push. This is the only way to finish work here — a plain
# `git push` does not bump, and a pre-push hook cannot do it either: git reads
# the refs before the hook runs, so a commit made inside the hook is never
# pushed.
#
#   ./tools/ship.sh "fix: what changed"   commit the work, bump, push
#   ./tools/ship.sh                       nothing to commit: bump, push
#
# Bump rule, from the commit subjects after the last commit that changed the
# version line: any `feat` -> minor, else any `fix` -> patch, else (docs, chore,
# refactor, test) no bump. The version line itself is the anchor, so running
# this twice never bumps twice.
set -euo pipefail
cd "$(dirname "$0")/.."

msg="${1:-}"
if [ -n "$(git status --porcelain)" ]; then
  if [ -z "$msg" ]; then
    echo "error: work tree is dirty, pass a conventional commit message" >&2
    echo "       ./tools/ship.sh \"fix: what changed\"" >&2
    exit 1
  fi
  git add -A
  git commit -q -m "$msg"
fi

anchor=$(git log -1 --format=%H -G'^version = ' -- src-tauri/Cargo.toml || true)
range="HEAD"
[ -n "$anchor" ] && range="$anchor..HEAD"
subjects=$(git log --format=%s "$range")

cur=$(sed -n 's/^version = "\(.*\)"$/\1/p' src-tauri/Cargo.toml | head -1)
[ -n "$cur" ] || { echo "error: no version line in src-tauri/Cargo.toml" >&2; exit 1; }

bump=""
if grep -qE '^feat' <<<"$subjects"; then
  bump=minor
elif grep -qE '^fix' <<<"$subjects"; then
  bump=patch
fi

if [ -z "$bump" ]; then
  echo "nothing user-visible since $cur: pushing without a bump"
else
  IFS=. read -r major minor patch <<<"$cur"
  if [ "$bump" = minor ]; then
    minor=$((minor + 1)); patch=0
  else
    patch=$((patch + 1))
  fi
  new="$major.$minor.$patch"
  echo "version $cur -> $new ($bump)"

  sed -i "s/^version = \"$cur\"$/version = \"$new\"/" src-tauri/Cargo.toml
  # tauri.conf.json declares the same version; tauri-build rejects a drift
  sed -i "s/\"version\": \"$cur\"/\"version\": \"$new\"/" src-tauri/tauri.conf.json
  grep -q "\"version\": \"$new\"" src-tauri/tauri.conf.json ||
    { echo "error: src-tauri/tauri.conf.json was not updated" >&2; exit 1; }
  # Cargo.lock keeps the package's own version too; `cargo metadata` would fix it
  # but it silently accepts a broken lock, so patch the one block and verify that.
  awk -v v="$new" '
    /^name = "isel-to-gcode"$/ { hit = 1 }
    hit && /^version = / { sub(/".*"/, "\"" v "\""); hit = 0 }
    { print }
  ' src-tauri/Cargo.lock > src-tauri/Cargo.lock.tmp
  mv src-tauri/Cargo.lock.tmp src-tauri/Cargo.lock
  grep -A1 '^name = "isel-to-gcode"$' src-tauri/Cargo.lock | grep -q "\"$new\"" ||
    { echo "error: src-tauri/Cargo.lock was not updated" >&2; exit 1; }
  sed -i "s|badge/version-[0-9.]*-blue|badge/version-$new-blue|" README.md

  git add -A
  git commit -q -m "chore: version $new"
fi

git push origin HEAD
v=$(sed -n 's/^version = "\(.*\)"$/\1/p' src-tauri/Cargo.toml | head -1)
echo "pushed. the Release workflow publishes v$v and attaches the exe automatically"
