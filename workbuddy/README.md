# WorkBuddy 安装包

本目录是 **WorkBuddy 适配层**，不是第二份技能源码。技能正文、脚本和模板仍以
`.agents/skills/` 为唯一真源；`scripts/build_workbuddy.py` 只生成可安装副本。

## 安装包内容

| 文件 | 用途 |
|---|---|
| `hetong-craft.skill` | Word 深读、红线/批注、起草、修订、Office 工具 |
| `hetong-review.skill` | 深度审核、版本对比、保护性条款、关闭核验 |
| `hetong-server.skill` | 服务器采购/销售合同 A01—A04/B01—B11 与暗改检测 |
| `hetong.skill` | **单包上传版**，内含三套技能，推荐给 WorkBuddy 图形界面 |
| `hetong-workbuddy.zip` | 三套技能同级合并包，适合直接复制到 `~/.workbuddy/skills` |

`hetong.skill` 把三套技能放在 `hetong/bundled/` 内，单次上传即可；三套单技能包必须安装为**同级目录**，因为 `hetong-review` 和 `hetong-server` 会引用 `../hetong-craft/scripts/`。

## 下载

开发预览版：

https://github.com/mmhsdxwdxx/hetong-skills/releases/tag/v1.3.0-dev.1

下载 `hetong.skill` 即可。正式版仍为 `v1.2.0`；WorkBuddy 预览包属于开发版，不替换正式版。

## 方式一：WorkBuddy 图形界面（推荐）

1. 打开 WorkBuddy → 左侧「专家·技能·连接器」→「技能」；
2. 「添加技能」→「上传技能」；
3. 上传 **`hetong.skill`** 一个文件即可；
4. 重启或重新打开技能页；
5. 在对话里直接说：

```text
用 hetong 审核这份合同，我方是乙方，先只读，不改原件。
```

如果你更习惯分开管理，也可以依次上传 `hetong-craft.skill`、`hetong-review.skill`、
`hetong-server.skill`，但三者必须同时安装。

## 方式二：直接复制（终端）

macOS / Linux，单包：

```bash
mkdir -p ~/.workbuddy/skills && unzip -q -o hetong.skill -d ~/.workbuddy/skills
```

如果使用三套单包，复制合并包：

```bash
mkdir -p ~/.workbuddy/skills && unzip -q -o hetong-workbuddy.zip -d ~/.workbuddy/skills
```

Windows PowerShell，单包：

```powershell
New-Item -ItemType Directory -Force "$HOME\.workbuddy\skills" | Out-Null
tar -xf .\hetong.skill -C "$HOME\.workbuddy\skills"
```

Windows 没有 `tar` 时，先把 `.skill` 复制成 `.zip`，再 `Expand-Archive`。

安装只新增 `hetong`（或三套同名技能）目录，不迁移、不删除你原有的 WorkBuddy 技能和资产。
如果本机已有旧版 `hetong-server`，先改名备份，否则会覆盖：

```bash
mv ~/.workbuddy/skills/hetong-server ~/.workbuddy/skills/hetong-server.backup
```

然后重启 WorkBuddy。


## 运行环境

- 三套技能的合同文本核心脚本只依赖 **Python 3.9+ 标准库**，不需要 pip 安装；
- WorkBuddy 自带的 Python 可用时直接运行；Windows 可用 `python` 或 `py -3`；
- Office 页面渲染/PPTX 精修需要额外安装 LibreOffice，但没有它不影响 DOCX 深读、红线、审核和暗改检测；
- 真实合同只在本地处理；不要把合同内容上传到公开服务。

## 重新生成

在公开仓库根目录运行：

```bash
python3 scripts/build_workbuddy.py
```

生成物在 `dist/`，不进入 Git；发布时作为 Release 附件上传。

## 验证边界

生成包已用 WorkBuddy 自带 `quick_validate.py` 校验，并检查三套技能同级安装后的脚本入口。
尚未在两台真实 Win11 + WorkBuddy 客户端上完成上传、Word/WPS 打开和 Office 渲染的实机验收；
正式发布前应补做这一层验证。
