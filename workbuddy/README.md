# WorkBuddy开发版安装包

本目录是适配说明，不维护第二份技能源码。scripts/build_workbuddy.py从.agents/skills生成包：

| 文件 | 内容 |
|---|---|
| hetong-craft.skill | 统一入口、甲乙方条款、Word、受理与时间轴 |
| hetong-review.skill | 深度审核、版本对比与关闭核验 |
| hetong-fast-xqy-v1.skill | XQY八项快审及双Word交付 |
| hetong.skill | 三套技能置于hetong/bundled的单包上传版 |
| hetong-workbuddy.zip | 三套技能的同级目录合并包 |

```bash
python3 scripts/build_workbuddy.py
```

生成物和SHA-256清单在dist，不进入Git。当前develop为2.0.0-dev.2，尚未发布对应Release附件；历史v1.3.0-dev.1下载不代表当前源码。

在WorkBuddy技能界面上传新生成的hetong.skill，或将合并包解压到宿主规定的skills目录。使用单技能包时须保持三套目录同级，以满足相对引用。升级前备份旧安装；独立hetong-server已撤并，旧入口应另行处理，不让两套架构混用。

安装完成后新开会话；本流程不刷新或重启已有线程。真实合同留在本地。核心读取与审核脚本以Python标准库为主；快审Word生成需要python-docx，Office页面渲染另需LibreOffice及字体。

构建和归档检查只能证明包结构、引用与源文件一致性。WorkBuddy界面上传、模型调用及Word/WPS实机行为须另行验证，不能以ZIP生成成功代替。
