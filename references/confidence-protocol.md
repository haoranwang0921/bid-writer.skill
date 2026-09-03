# 环节二：置信度判定与提问协议

## 槽位（slot）定义

对环节一切片产物 `template.docx` 直接扫描，槽位 = 占位实例：

| type | id 格式 | label 取法 |
|---|---|---|
| paragraph | `P{段号}#{该段第几个占位}` | 占位符前 18 字窗口归一化 |
| table_cell | `T{表号}:R{行}C{列}` | 表头行同列文本，缺则用单元格现文本 |

占位识别正则（合并识别）：`_{2,}`、`＿{2,}`、`×{3,}`、`【…】`、`（ ）`、`( )`、**冒号空白槽** `[：:]+≥2空格(?=（提示语/年/元/天/日/%/行尾)`——覆盖「致：      （招标人名称）」「日期：   年  月  日」类 Word 空格填空。

## 资料源（--sources）

任意 JSON 扁平化为 `归一化键 → [{value, source}]`。企业画像 schema：

```jsonc
{
  "basic": {"name":"", "uscc":"", "legal_representative":"", "registered_capital":"",
             "established":"", "address":"", "contact":{"phone":"","email":""}, "bank":{}},
  "qualifications": {"营业执照":{"no":"","valid_until":""}, "安全生产许可证":{…}, …},
  "performance": [{"project":"","client":"","contract_date":"","amount":"","evidence":""}],
  "personnel": {"项目经理":[{"name":"","certs":[]}], "技术负责人":[…]},
  "equipment": [], "finance": {}, "declared_facts": {"consortium": false, "…": ""}
}
```

当前工作区实例：`AI标书撰写/企业画像_雅合科技.json`（**草稿态**：文本层可提取字段已填，证照编号/签章人等待扫描件核对，字段值里的「待确认」会自然落入 low 档提问）。

## 置信度判定规则

槽位 label 与源键 norm 后双向包含匹配（≥2 字）：

- **high**：候选值唯一（同值多源合并）→ `status=auto`，直接进 fills；
- **medium**：2–4 个不同候选值 → `status=ask`，提问列出候选+依据来源；
- **low**：零命中 → `status=ask`，选项为「自行填写」（agent 可给建议选项，但建议不落 answers 即不生效）。

> **语义复核义务**：label 匹配是机械初筛，提问前 agent 必须剔除形似而实非的命中（如"地址"槽泛匹配到业绩地址）、把"多套合理口径"（含税/不含税、大写/小写栏）升为 medium。**脚本保证不漏，人保证不错。**

## 提问协议（AskUserQuestion 映射）

questions.json 每条 → AskUserQuestion 一题：
- `question`：`「{context 所在小节}」中的「{label}」如何填写？`
- `options`：candidates 前 3 项（label=值，description=依据来源）；low 题靠系统 Other 自填
- 每批 ≤4 题；全部答复汇总进 answers.json 后才执行 apply。
- **W2**：answers.json 是唯一写入凭据；未答槽位在成稿中保持原占位，交付前必须清零。

## 接口契约

```jsonc
// fill_plan.json（scan 输出）
{"meta":{"template","generated_at","stats":{"total","high","medium","low"}},
 "slots":[{"id","type","label","current","context","candidates":[{"value","source"}],"confidence","status"}]}
// questions.json（scan 副产物，供提问）
{"questions":[{"id","label","current","context","confidence","options":[…],"basis"}]}
// answers.json（用户确认后 agent 整理）
{"answers": {"P148#0": "青岛雅合科技发展有限公司", "T3:R1C2": "2026-09-01"}}
// fills.json（apply 输出）
{"fills": {"<slot_id>": "<值>"}, "unconfirmed": 7}
```

## kb_assist.py（环节二可选增强，本地知识库检索）

**作用**：对 `status=ask` 的槽位（medium/low），调用本地检索服务（`127.0.0.1:8765`）获取参考知识片段，写入槽位的 `kb_hint` 字段，供提问关口展示给用户作为判断依据。

**用法**：
```bash
python scripts/kb_assist.py fill_plan.json -o fill_plan_kb.json
# 服务不可达时静默跳过，输出与输入相同（不阻断流程）
```

**槽位新增字段**（仅 ask 档）：
```jsonc
// fill_plan_kb.json 中 ask 档槽位示例
{"id": "P42#0", "label": "接地电阻要求", "confidence": "low", "status": "ask",
 "candidates": [],
 "kb_hint": "【知识库引用 1】相似度 0.80 | 来源：知识库/02-技术参数库/排流地床技术参数.md…\n辅助阳极地床接地电阻 <4Ω…"}
```

**使用纪律**（与 W2 一致）：
- `kb_hint` 是**背景参考**，帮助判断口径，不作为候选值、不绕过确认；
- medium 槽位：`kb_hint` 解释候选值来源背景（如"填包料配方：石膏粉75%/硫酸钠5%/膨润土20%"）；
- low 槽位：`kb_hint` 提供填写方向，最终值仍须用户确认写入 `answers.json`；
- `meta.kb_assist.status`：`done`（检索完成）/ `skipped`（服务不可达）。

## fill_docx.py（就地写回）

- 段落槽：run 级替换优先（保字体格式）；占位跨 run 时合并到首 run 重写（会损失该句内局部格式，替换后肉眼复查）。
- 表格槽：定位 (表,行,列) 单元格内首个占位替换。
- 退出码 3 = 剩余占位 >0 → 回提问关口；**禁止**用"看起来合理"的内容凑数清零。
- 写回后文件即投标草稿：叙述章节在此文件上续写，**模板既有节只可填不可删**（verify C 轨道核对 template.json headings 全覆盖）。
