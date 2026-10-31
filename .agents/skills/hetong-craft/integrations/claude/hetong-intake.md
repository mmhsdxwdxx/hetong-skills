---
description: 用Claude原生单选和复选收集合同受理范围，保存真实提交回执。
argument-hint: 合同原件路径与本次已明确的处理要求
---

执行 hetong-craft 的受理流程，读取该技能 references/host-intake.md，使用 scripts/hetong_intake.py。用户输入：$ARGUMENTS

锁定源文件，把用户已明确的信息及其原话交给 init；未知项不填。已有有效回执仅补缺失项。合同任务直接使用公开CLI，不导入或猜测内部Python函数。

执行 plan --host claude --out 问题批次.json，读取输出；questions为空则读取status返回，否则把 arguments 原样传给主代理的真实 AskUserQuestion。甲方前置题提供A—D四种保交付/资金共管组合，随后复选其余类别；乙方直接复选当前有条目的类别及“本次不用乙方预设条款”；必须保持各题 multiSelect 值。不要调用 collect 打开网页，不要把打印工具参数当作实际弹问。

将实际工具结果原样保存到本任务产出，执行 accept --host claude；未完成且确有下一批问题时继续 plan → AskUserQuestion → accept。不得预填或伪造答案。取消、跳过、超时或自由文本改意时保留待选择，不自动重弹。当前宿主没有原生工具时明确说明，用户要求原生则保留待选择。

收到 submitted 后读取 status，并按其处理方式返回主任务。两题只确定预设条款范围，不确认金额、交期、担保、共管资金安排或正文写入授权。只执行受理，不修改合同。
