# contract-review-agent

> 来源分类：**原创/AI打磨** ｜ 导出批次：published

合同审查智能体（对应 C002 FR-07）。输入合同文本(txt/md/docx导出文本)，AI 逐条拆解条款、识别风险点（合法性/合规性/权利义务失衡/缺失条款/表述歧义），按高/中/低分级并给出修改建议与法条依据，输出条款级批注版审查报告(Markdown/DOCX)。触发词：合同审查、审合同、合同风险、合同把关、审查报告、合同风险点、review contract。

## 安装

把本文件夹整体复制到 WorkBuddy 技能目录：

```bash
cp -r . ~/.workbuddy/skills/contract-review-agent        # 用户级
# 或
cp -r . <项目>/.workbuddy/skills/contract-review-agent   # 项目级
```

重启/刷新 WorkBuddy 后即可在对话中触发。

## 说明

- 本技能从本地 WorkBuddy 环境导出，**所有真实密钥已脱敏为占位符**，使用前请配置你自己的 API Key。
- 若来自技能市场（文件夹名以 `__skillhub` 结尾），版权归原作者，请遵守其许可证。
