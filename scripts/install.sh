#!/bin/sh
set -eu

if [ "$#" -ne 1 ]; then
    printf '用法: %s <目标项目目录>\n' "$0" >&2
    exit 2
fi

TARGET=$1
ROOT=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)

if [ ! -d "$TARGET" ]; then
    printf '目标目录不存在: %s\n' "$TARGET" >&2
    exit 2
fi
if ! command -v rsync >/dev/null 2>&1; then
    printf '需要 rsync 才能安装；也可手动复制 .agents/skills/。\n' >&2
    exit 2
fi

mkdir -p "$TARGET/.agents/skills"
rsync -a "$ROOT/.agents/skills/" "$TARGET/.agents/skills/"
printf '已安装到 %s/.agents/skills/\n' "$TARGET"
printf '请从目标项目根目录启动 Pi/Codex/DSH，并在需要时重载技能。\n'
