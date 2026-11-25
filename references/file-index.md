# 分发版文件清单与能力对照

> 本技能采用开放平台规定的两级目录：技能根目录 ＋ `references/`／`scripts/`／`templates/` 三个一级子目录。
> 所有路径相对于技能根目录，无绝对路径、无深层嵌套、无外部服务依赖。

## 目录结构

```
new-cuoti-pipeline/
├── SKILL.md
├── README.md
├── references/    规则、口径、经验库、模板
├── scripts/       可执行脚本（Python 标准库 + python-docx）
└── templates/     示例卷结构化数据
```

## 能力对照（与全流程一一对应）

| 流程层 | 能力 | 承载文件 |
|---|---|---|
| L1 识别 | 三色分离、按题裁切、转写铁律、分数反算自洽 | `SKILL.md` L1、`references/三步口径.md`、`references/识别经验库.md` |
| L1.5 题面审查 | 六维 × 五视角、VALID/SUSPECT/INVALID 三态、提示文案 | `references/题面有效性审查.md` |
| L2.0 答案交叉论证 | 六路独立求解、红队反证、收敛裁决、正确率口径 | `references/答案交叉论证.md`、`scripts/verify_answers.py` |
| L2 判定 | 四维判定、客观题自动判、主观题不自判 | `SKILL.md` L2、`references/三步口径.md` |
| L3 确认 | 八列判定表、四张存疑清单、确认后入库 | `SKILL.md` L3 |
| L4 归因 | 五类 A–E 错因、结构化错题库记录 | `SKILL.md` L4 |
| L5 变式复习 | 变式三招、机器验证、艾宾浩斯排期 | `SKILL.md` L5、`scripts/verify_answers.py` |
| L6 呈现 | 家长版＋学生版双份 Word、三色审计、隐私扫描、v2.0 门禁 | `scripts/build_books.py`、`scripts/docx_kit.py`、`references/白话改写规则.md`、`references/数据契约.md`、`references/验收清单.md` |

## 打包说明（相对本地定制版）

分发版**保留全部主流程能力与判定口径**；移除的是不可移植的本机专属部分：

1. 深层嵌套的旧引擎目录（含开发测试文件、Web App 模板、变更日志）——平台仅允许两级目录。
2. 本机错题库系统及其固定数据目录——改由模型按 L4/L5 口径直接产出结构化结果，可对接任意本地存储。
3. 依赖外部账号的可选语音讲解模块——主流程本就不需要，移除后不影响任何功能。
4. 所有本机绝对路径、用户目录名、地区与考生专属配置——改为相对路径与可填写模板。

能力等价性由 `scripts/build_books.py` 的回归测试验证：示例卷数据可直接生成双份 Word，三色审计、隐私扫描与 v2.0 核验字段自检均正常。
