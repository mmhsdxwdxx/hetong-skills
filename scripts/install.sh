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
printf '请从目标项目根目录新开会话使用新版；本安装器不会刷新已有会话。\n'
printf '旧 hetong-server 已撤并；如目标中仍有旧入口，请先自行备份处理，安装器不会删除它。\n'
