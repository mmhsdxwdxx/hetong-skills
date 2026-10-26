# hetong-skills

面向中文商务合同的技能包，供 Pi Agent、Codex 和其他支持 Agent Skills 的宿主使用。
本仓库是公开发布包，**不包含真实合同、客户材料或审核成果**。

## 包含内容

```text
.agents/skills/
├── hetong-craft/     # 起草、修订、润色、完善、审核统一入口；Word 红线/批注、Word 深读、Office 工具
├── hetong-review/    # 深度审核、跨合同勾稽、版本对比、关闭核验与保护性条款
└── hetong-server/    # 服务器采购/销售合同专用 profile；乙方 B01—B11、甲方 A01—A04、两版暗改检测
```

核心方法、模板、脚本和虚构测试材料均在技能目录内。`hetong-server` 的公开版本已移除
真实交易排期、历史金额/台数、源文档指纹和具体署名，仅保留可复用的方法结构。乙方视角
B01—B11 与甲方视角 A01—A04 分列；A01—A04 覆盖保函受益人/生效、BIS 与政府禁令、保函
有效期与提取、资金共管账户及 U 盾权限。`hetong-server/scripts/hidden_change.py` 可对
前后两版近似 DOCX 检查绕过修订记录的直接改文，并区分留痕改动、历史减少和非文本变化。

## 安装

把包复制到目标项目的 `.agents/skills/`：

```bash
git clone https://github.com/mmhsdxwdxx/hetong-skills.git
cd hetong-skills
./scripts/install.sh /path/to/your-project
```

也可以手动复制：

```bash
mkdir -p /path/to/your-project/.agents/skills
cp -R .agents/skills/. /path/to/your-project/.agents/skills/
```

安装脚本使用 `rsync -a` 覆盖同名技能文件，但**不会删除**目标项目中的其他技能或文件。
如果目标项目已有本地修改，请先备份或使用版本控制审查差异。

### Windows 11

Windows 原生使用 PowerShell，不需要 WSL：

```powershell
git clone --branch develop https://github.com/mmhsdxwdxx/hetong-skills.git
cd hetong-skills
.\scripts\install.ps1 -Target "$HOME\Projects\hetong-test"
```

也可以直接运行 `scripts\install.cmd`。Windows 安装器使用系统自带的 `robocopy`，不会删除
目标目录中的额外文件。Office 运行环境：

```powershell
py -3.12 .agents\skills\hetong-craft\scripts\office_kit.py setup
py -3.12 .agents\skills\hetong-craft\scripts\office_kit.py doctor
```

隔离环境在 `%LOCALAPPDATA%\contact-office\venv`；LibreOffice 默认从
`C:\Program Files\LibreOffice\program\soffice.exe` 发现。Windows 使用系统字体，
不生成 Unix fontconfig 配置。

从目标项目根目录启动宿主，技能才会通过项目 `.agents/skills/` 被发现。Pi 可显式使用：

```text
/skill:hetong-craft
/skill:hetong-review
/skill:hetong-server
```

通用合同使用 `hetong-craft` / `hetong-review`；只有服务器设备采购/销售合同才加载
`hetong-server`。我方为甲方时使用 A01—A04，为乙方时使用 B01—B11。自然语言说明技能名称也可以。

## 更新

```bash
cd /path/to/hetong-skills
git pull --ff-only
./scripts/install.sh /path/to/your-project
```

更新后新开宿主会话或执行相应重载操作。安装脚本不删除目标项目中的额外文件，便于保留
本地扩展；但会覆盖同名技能文件。当前版本见 `VERSION` 与 GitHub Release。

## 环境

- Python 3.9+；合同文本核心脚本以标准库为主。
- Office 深读、PPTX 精修和无头渲染可通过隔离环境安装依赖：
  `python3 .agents/skills/hetong-craft/scripts/office_kit.py setup`。
- LibreOffice 仅用于转换和页面渲染，不是合同写入的依赖；缺失时 Office 预览不可用，
  纯 DOCX 阅读和红线工具仍可运行。
- Word/WPS 的最终显示、自动编号、复杂表格和分页仍应在目标客户端复核。

## 验证

```bash
python3 .agents/skills/hetong-craft/scripts/test_hetong_redline.py
python3 .agents/skills/hetong-craft/scripts/test_word_reader.py
python3 .agents/skills/hetong-craft/scripts/test_office_kit.py
python3 .agents/skills/hetong-craft/scripts/test_review_audit.py
python3 .agents/skills/hetong-craft/scripts/test_platform_compat.py
python3 .agents/skills/hetong-server/scripts/test_hidden_change.py
python3 .agents/skills/hetong-review/scripts/test_review_kit.py
```

详细验证记录见 `.agents/skills/VALIDATION.md`。

## 安全边界

- 默认只读原件；不自动接受或拒绝修订，不自动删除批注。
- 修订和批注写入必须显式授权，并应保留原件、计划和验证记录。
- 不把真实合同、交易信息或客户材料上传到公开服务。
- 本仓库的示例、数字和虚构测试材料仅用于演示，不构成法律意见。
- 法律结论需要结合现行法、交易背景和用户确认另行核验。

## 许可证

MIT。详见 [LICENSE](LICENSE)。
