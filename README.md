# bid-writer

模板截取式投标文件撰写 Skill —— 面向 AI Agent 的投标 / 响应文件编写工作流。

> 与 bid-studio（全流程投标工作台：决策 → 解析 → 编制 → 校验 → 交付）互补。bid-writer 专注 **「模板驱动撰写 + 量化回归」**：模板只能来自招标文件原文，没有来源依据的内容一个字都不能写。

## 设计铁律

| 规则 | 内容 |
|---|---|
| W1 | 模板唯一来源是招标文件；定位失败必问用户，不得自造格式 |
| W2 | 只允许机械精确 high、经证据关卡放行的模型 high、或用户确认值写入；medium/low 必须确认 |
| W3 | 金额、资质编号、日期只信企业画像、当前招标事实与已确认事实，冲突即问 |
| W4 | verify FAIL 或剩余占位 >0 → 禁止交付 |
| W5 | 回归差异不得静默忽略，每轮 diff 报告归档 |

## 流水线

```
节点 0 ─ 招标文件解析 ──────────────► tender.json（常量/评分项/废标红线/时间节点）
环节一 ─ 模板截取 ──────────────────► template.docx（原样切片，继承全部格式）
环节二 ─ 模型检索 + 证据化填空 ─────► 投标文件草稿.docx（就地写回，格式不重建）
环节三 ─ 交付机检（四轨道）─────────► verify_report.md（FAIL 禁止交付）
环节四 ─ diff 回归（对标人工成品）──► diff_report.md + 通过率
```

| 阶段 | 脚本 | 输入 → 产出 | 强制关卡 |
|---|---|---|---|
| 节点 0 | `parse_tender.py` | 招标文件 docx → `tender.json` + 项目常量 | 空解析退出码 3，停下核对 |
| 环节一 | `extract_template.py` | 招标文件 → `template.docx` + `template.json` | 起止定位失败 / 必备件缺失 → 退出码 3，必须向用户确认 |
| 环节二 | `fill_plan.py` → `smart_fill.py` → 人机确认 → `fill_docx.py` | template + 数据库/知识库 → 带证据草稿 | 模型值须经证据关卡；未决槽位不写入 |
| 环节二附 | `write_narrative.py` | 叙述章节素材检索（本地知识库） | 数字/参数须带 `source_path` 溯源 |
| 环节三 | `verify.py` | 草稿 + template/quotes/tender → 校验报告 | A 占位残留=0 / B 金额一致 / C 模板节不删 / D 红线逐条响应 |
| 环节四 | `diff_report.py` | 人工成品 ↔ 新结果 → 差异清单 + 通过率 | 存在未达标节 → 退出码 4，修后重跑 |

## 目录结构

```
bid-writer/
├── SKILL.md                       # skill 主文档（工作流 + 实战踩坑记录）
├── references/
│   ├── template-schema.md         # 环节一：模板截取实现逻辑与元数据契约
│   ├── confidence-protocol.md     # 环节二：置信度判定与提问协议
│   ├── smart-fill-protocol.md     # 环节二：模型检索、证据结构、自动放行与审计协议
│   ├── diff-interface.md          # 环节四：diff 算法与报告格式
│   └── redline-checklist.md       # 按采购类型的必备模板/要件清单
└── scripts/
    ├── parse_tender.py            # 节点 0：招标结构解析
    ├── extract_template.py        # 环节一：起止定位 + 原样切片
    ├── fill_plan.py               # 环节二：槽位扫描 + 置信度分级 + 提问单（精确键=high，子串=medium）
    ├── smart_fill.py              # 环节二：生成模型检索任务 + 验证模型证据决策
    ├── fill_docx.py               # 环节二：fills 就地写回切片副本（保 rPr 下划线/字号/字体，跨 run 安全）
    ├── write_narrative.py         # 环节二续：叙述章节撰写（KB 检索 → 草稿）
    ├── insert_images.py           # 环节二附：从成交响应文件抽图按小节插入（标题锚点后居中图 6in）
    ├── verify.py                  # 环节三：交付机检（四轨道，缺参→SKIP 显式披露）
    ├── diff_report.py             # 环节四：diff 回归
    ├── kb_assist.py               # 可选：知识库检索辅助（附 kb_hint）
    └── _common.py                 # 共享：金额解析 / 归一化 / 原子写 / 路径守卫 / 豁免表
└── tests/                         # pytest 测试集（基础逻辑 / 槽位 / 智能证据）+ 冒烟脚本
```

## 快速开始

```bash
# 0) 解析招标文件（常量/评分项/废标红线/时间节点）
python scripts/parse_tender.py 招标文件.docx -o tender.json -k 项目常量.json --with-timeline

# 1) 截取模板（任何撰写动作之前必做；找不到模板就问，不许编模板）
python scripts/extract_template.py 招标文件.docx -o template.docx -m template.json \
    --require 投标函,报价表,偏离表,资格审查

# 2) 扫描槽位 → 模型查询数据库 → 证据关卡 → 人机确认 → 写回
python scripts/fill_plan.py scan template.docx --sources 企业画像.json,项目常量.json -o fill_plan.json --questions questions.json
# data_sources.json 是可选的数据源登记表；没有它时使用 fill_plan.meta.source_catalog
python scripts/smart_fill.py prepare fill_plan.json -o model_tasks.json
# 如使用额外的 SQLite/HTTP 数据源目录，则在上条命令后加：--catalog data_sources.json
# 运行 Skill 的 agent 读取 model_tasks.json，使用只读数据工具检索并生成 model_decisions.json
python scripts/smart_fill.py validate fill_plan.json model_decisions.json -o smart_fill.json --questions smart_questions.json
# 若 prepare 使用了外部目录，validate 必须使用同一目录：--catalog data_sources.json
python scripts/fill_plan.py apply template.docx fill_plan.json --smart smart_fill.json --answers answers.json -o fills.json
python scripts/fill_docx.py template.docx fills.json -o 投标文件草稿.docx

# 3) 交付机检（FAIL = 退出码 5，修复后重跑）
python scripts/verify.py 投标文件草稿.docx --template template.json \
    --quotes quotes.json --tender-parse tender.json -o verify_report.md

# 4) diff 回归（对标人工成品，未达标修流程重跑）
python scripts/diff_report.py 人工成品.docx 投标文件草稿.docx -o diff_report.md --threshold 0.85
```

Windows 下脚本一律以**绝对路径**调用（bash 中 `~` 与相对路径会被错误解析）。Python 依赖仅 `python-docx`。

## 置信度分级

| 置信度 | 判定 | 动作 |
|---|---|---|
| high | 画像键名与槽 label 归一化后【精确一致】（含短编号前缀如「（九）工期」）且值唯一 → auto | 直接填入，不问 |
| medium | 精确键多值冲突 / 仅子串（模糊）命中（如画像「名称」键 vs 槽「项目名称」） | 列出候选项 + 依据，向用户确认后才写入（**子串命中不得判 high 防串值**） |
| low | 无任何来源 | 必须提问（含「自行填写」选项），禁止臆测 |

槽位 = 切片 docx 中的占位：段落空白槽（`____`、`【】`、`（ ）`、冒号空白）与表格占位单元格。**模板原生格式标题豁免表**（如「投标函【工程量清单计价】」）见 `_common.EXEMPT_PH`，新项目遇新格式词在此追加。槽位 id 稳定格式：段落 `P{段号}#{第几个占位}`、表格 `T{表号}:R{行}C{列}`。

## 模型检索填空

模型读取每个槽位的 label、上下文和现有候选，自主查询只读企业库、项目库、招标解析结果或知识库。模型输出不能直接写 Word，必须先经过 `smart_fill.py validate`：只有唯一、可定位、来源已登记且来自 `structured` / `tender` / `confirmed` 的事实可自动填写；历史文档、向量知识库、冲突值和缺定位结果全部降级提问。完整数据源目录、决策 JSON 与审计契约见 `references/smart-fill-protocol.md`。

`fills.json.audit` 记录每个值来自 `human_answer`、`model_evidence` 还是 `deterministic_auto`，写入优先级为人工 > 模型证据 > 机械 high。

`data_sources.json` 不随 Skill 打包，因为企业库和项目库属于项目数据。它只登记 `id`、`kind`、`location`、`source_type` 和 `access: read_only`；完整格式见 [`references/smart-fill-protocol.md`](references/smart-fill-protocol.md)。Skill 不直接绑定某一家模型或数据库，运行时 agent 负责使用当前环境可用的只读连接器。

## 本地知识库辅助（降级方案）

环节二的知识库检索由本地 ChromaDB 服务提供（`127.0.0.1:8765`，向量模型 `BAAI/bge-small-zh-v1.5`），**数据不出本机、不依赖云端**：

- 脚本通道：`python scripts/kb_assist.py fill_plan.json -o fill_plan_kb.json`（为 ask 档槽位附 `kb_hint`）
- HTTP 通道：`POST /query`（续写叙述章节时直接取 `context_block`）

纪律：`kb_hint` / 检索片段只是背景参考；在模型智能模式中可以整理成 medium 候选，但不能仅凭向量库结果自动填写。引用的数字/参数必须带 `source_path` 溯源；服务不可达时静默跳过，不阻断流程。

## 实战验证

2026-09 淄博项目（市政工程，施工总承包类响应文件）全流程验证，踩坑与解法已沉淀进 `SKILL.md`，要点包括：

- 模板起止定位：自动命中目录页/引言短句时判 MISS，正确解法是 `--start 关键词` + 显式 `--end 下一章标题原文`
- 知识库检索：历史成交响应文件可用于帮助模型定位企业要素，但需核验并固化为当前结构化事实后才允许自动填入
- 投标函等格式标题（如「投标函【工程量清单计价】」）会被误判为占位槽，填空脚本须排除
- 报价「留占位」交付时，verify A 项残留属预期 FAIL → 输出《补录清单》并标注不可正式递交
- 红线扫描措辞须与招标原文逐字对齐（大小写、含「的」），可做到 20/20 全命中

## 环境要求

- Python 3.10+
- 运行依赖：`python-docx`
- 测试依赖：`pytest`（49 个用例 + 冒烟脚本）
- 企业画像 JSON 保存在项目工作区，skill 内不存放任何真实企业数据

## 版本

- v0.4.0（2026-09-04）：模型检索填空。新增任务生成、模型决策证据关卡、数据源目录和填写审计；可信唯一事实自动填，冲突/文档/向量证据降级提问
- v0.3.2（2026-09-04）：下划线格式保真。fill_docx 跨 run 替换不再清空旧 run 重建，改为填入文字落在首 run（保留 rPr）+ post 段追加带源 rPr 新 run；新增 test_underline.py 5 用例
- v0.3.1（2026-09-04）：质量加固。fill_plan 精确键=high、子串降 medium；verify SKIP 披露 + D 红线短行判定；补 insert_images.py；_common.norm/cn_to_num/parse_amount bugfix；tests/ 33 用例 + 冒烟
- v0.3.0（2026-09-01）：四环节流水线定型；淄博项目实战踩坑记录；知识库直填 / 图片插入流程沉淀
- 首版发布于 2026-09-03
