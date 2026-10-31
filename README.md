# hetong-skills

中文商务合同技能包，支持项目级Agent Skills。此仓库仅包含通用方法、经授权的条款库及虚构测试材料，不包含真实合同、客户资料或审核成果。

本分支是 **2.0.0-dev.1 开发版**。稳定安装固定到已发布标签，不把develop或main的最新提交当作稳定版本。

| 入口 | 用途 |
|---|---|
| hetong-craft | 五模式统一入口、甲乙方条款、受理、Word与Office工具 |
| hetong-review | 深度审核、版本对比、关闭核验、保护性条款 |
| hetong-fast-xqy-v1 | XQY八项快审：详细Word报告＋原稿批注副本 |

独立hetong-server已撤并至craft。甲方A01—A24来源整合为20项，乙方B01—B11按类别使用；先锁定真实交易角色与本次范围，不将库中参数当成双方已定条件。旧安装升级前先备份；安装器不自动删除旧server目录，以免损失本地修改。

## 安装与新会话

开发试用：

```bash
git clone --branch develop https://github.com/mmhsdxwdxx/hetong-skills.git
cd hetong-skills
./scripts/install.sh /path/to/your-project
```

Windows PowerShell：

```powershell
git clone --branch develop https://github.com/mmhsdxwdxx/hetong-skills.git
cd hetong-skills
.\scripts\install.ps1 -Target "$HOME\Projects\hetong-test"
```

从目标项目根目录**新开会话**使用新版。安装脚本复制`.agents/skills/`，不发送刷新、重载或重启指令给既有线程，也不切换模型。既有会话若主动重新读取磁盘文件仍可能读到新内容；文件更新不等于为所有旧会话创建版本隔离。

本仓库另含`.claude/skills/`、`.claude/agents/`、`.claude/commands/`及`.pi/extensions/`相对链接，用于直接从此检出启动的本地项目。安装到其他项目时，按需将对应链接或链接目标复制至相同相对位置；Windows或不支持软链接的宿主应复制目标文件，不能使用链接文本。不要覆盖已有同名本地配置。宿主加载以实际探针为准。

```text
请用hetong-craft审核这份采购合同，我方为甲方，先只审不改。
xqy快审，站乙方立场，合同位置：……。
请用统一时间轴检查付款、验收、保函与发票是否衔接。
```

## 时间轴与批注

[履约时间轴规范](.agents/skills/hetong-craft/references/contract-timeline.md)固定日期栏及三栏正文：我方现金流及付款义务、交付与担保、发票（万元）。A4版心159.2mm，四列22／52／50／35.2mm，每要点独立换行，无合并标题行。

[Markdown虚构示例](.agents/skills/hetong-craft/assets/contract-timeline-example.md)独立可读；[HTML模板](.agents/skills/hetong-craft/assets/contract-timeline-template.html)不依赖网络，提供稳定列宽与打印版式。应付、已付、待退和资金占用分开，正常履行与违约分支分开。完整图放分析报告或获准附件；批注跟随实际修改或具体问题，不因命中固定条款机械新增。

## 环境与验证

核心合同读取、红线、审阅及类别工具以Python 3.9+标准库为主。快审Word生成需要python-docx；Office转换和页面预览需要LibreOffice及相应字体。使用已有Office环境，或按需运行：

```bash
python3 .agents/skills/hetong-craft/scripts/office_kit.py setup
python3 .agents/skills/hetong-craft/scripts/office_kit.py doctor --full
```

测试命令见[技能维护说明](.agents/skills/README.md)，公开验证范围见[VALIDATION](.agents/skills/VALIDATION.md)。CI运行核心回归、快审生成、安装及打包检查；无头预览与Word/WPS实机往返分别报告，不能互相替代。

## WorkBuddy与分发

运行`python3 scripts/build_workbuddy.py`可从当前源码生成新版hetong.skill、三套单包和合并包，详见[WorkBuddy说明](workbuddy/README.md)。dist不入Git，本次仅同步develop源码，未发布新的Release附件；历史v1.3.0-dev.1下载仍是旧架构。

## 安全与授权

默认只读原件；修订与批注必须在授权范围内另存并保留历史。商业参数以明确确认值为准；法律结论核对现行正式法源。不要将真实合同、账户或客户材料放入公共仓库。示例仅用于演示，不构成交易合意或具体法律意见。

MIT许可证，见[LICENSE](LICENSE)。
