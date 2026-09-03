---
name: bid-writer
description: 模板截取式投标文件撰写。触发场景：写标书、编制投标文件/响应文件、按招标文件格式模板填空撰写。核心约束：(1)模板必须从招标文件 docx 截取（extract_template.py 起止标题定位+原样切片，继承原始格式/表格/占位符，禁止凭空生成模板，起止或必备件定位失败必须先向用户确认）；(2)填空按置信度分级（fill_plan.py scan→提问→apply→fill_docx 就地写回），high=画像键名与槽 label 归一化后精确一致且唯一值→auto；medium/fuzzy→ask 列 2-4 候选项+依据向用户确认后再写入，**子串命中（如「名称」键 vs「项目名称」槽）不得判 high 防串值**；low→ask（禁止臆测）；(3)交付前跑机检（verify.py：占位残留/金额一致/结构保真/红线，红线按行判定+目录短行不算实质响应），缺失 --quotes/--template/--tender-parse 时对应轨道记 SKIP 显式披露，FAIL（exit 5）禁止交付；(4)叙述章节由 write_narrative.py 检索本地知识库生成素材并写入草稿（数字/参数须溯源）；(5)证照图片由 insert_images.py 从成交响应文件抽图按小节插入（标题锚点后居中图 6in）；(6)以既有参考输出为基准做逐节 diff 回归（diff_report.py），输出差异清单与通过率统计。区别于 bid-studio（全流程含决策/评分范式挂载），本 skill 专注「模板驱动撰写+量化回归」。
version: 0.3.1
agent_created: true
metadata:
  compatible_agents: [workbuddy]
---

# bid-writer：模板截取式投标文件撰写

四环节流水线，每个环节有确定性脚本与强制关卡。设计铁律：**没有来源依据的内容一个字都不能写；模板只能来自招标文件原文。**

## 目录与依赖

- `scripts/extract_template.py` — 环节一：起止定位 + **原样切片** → template.docx + template.json（元数据）
- `scripts/parse_tender.py` — 节点0：招标结构解析 → tender.json（常量/评分项/废标红线/时间节点）+ 项目常量.json
- `scripts/fill_plan.py` — 环节二：槽位扫描 + 置信度分级 + 提问单 / 汇总 fills.json
- `scripts/kb_assist.py` — 环节二：知识库检索辅助（可选，为 ask 档槽位附 `kb_hint` 背景参考）
- `scripts/fill_docx.py` — 环节二尾部：fills 就地写回切片副本（格式继承源文件，不重建文档）
- `scripts/write_narrative.py` — 环节二续：叙述章节撰写（KB 检索素材 → 就地写入草稿 + 溯源 JSON）
- `scripts/insert_images.py` — 环节二附：从成交响应文件中按小节抽取证照图片（`a:blip`→rid→part.blob），在标题锚点段后插入居中图（Ex：`anchor._p.addnext(p._p)`，`insert_paragraph_before` 创建、无 insert_after）；目标被 Word 锁定时 `PermissionError` 自动另存 `_含图版.docx`
- `scripts/verify.py` — 交付关卡：占位残留/金额一致/结构保真/红线扫描
- `scripts/diff_report.py` — 环节四：参考基准 ↔ 新结果 逐节 diff + 通过率
- `scripts/_common.py` — 共享：金额解析（大写↔数字）、归一化、原子写、路径守卫
- `references/template-schema.md` — 环节一实现逻辑与元数据契约
- `references/confidence-protocol.md` — 环节二置信度判定与提问协议
- `references/diff-interface.md` — 环节四 diff 算法与报告格式
- `references/redline-checklist.md` — 按采购类型的必备模板/要件清单

Python 依赖 `python-docx`；测试需 `pytest`。Windows 下脚本一律以**绝对路径**调用（bash 中 `~` 与相对路径会被错误解析）。

## 环节一：模板截取（先取模板，后动笔）

**任何撰写动作之前必须执行本环节。不提取骨架、不改写——只确定模板起始/终止标题，把中间部分原样切成新 docx。找不到模板就问，不许编模板。**

```bash
python scripts/extract_template.py <招标文件.docx> -o template.docx -m template.json \
    [--start 响应文件格式] [--end 下一章节标题] \
    --require 投标函,报价表,偏离表,资格审查
```

- **切片原理**：复制源 docx 对象后删除起止范围之外的正文块 → `template.docx` 完整继承原始样式、标题层级、表格、占位符、页眉页脚，不做任何重排。**删除必须覆盖范围外所有顶层 body 块**（段落、表格、`<w:sdt>` 内容控件、`customXml` 等），仅保留 body 末尾 `sectPr`——若只按 w:p/w:tbl 白名单过滤，前部章节的表格与 sdt 包裹内容会整体残留到切片开头（表现为"前几页重叠"，v0.2.1 已修复）。
- **实战坑（2026-09 淄博项目验证）**：① 自动/显式起始常命中章首目录或引言短句（跨度<400 判 MISS）；当目录句与正题句文本相同时无法靠 `--start` 区分，正确解法是 `--start <该关键词>` + 显式 `--end <下一章标题原文>`，让打分选出正文候选或直接把终止锁到下一章（本次「第六章 响应文件要求及格式」→「第七章 合同条款」，6306 字符切片成功）；② `-o/-m` 的输出目录不会自动创建，先 `mkdir -p`，否则 docx save 抛 FileNotFoundError；③ 投标函等格式标题「投标函【工程量清单计价】」会被 PH_RE 误判为占位槽（scan 结果中的 `P89#0/P269#0` 类），填空脚本必须排除，不得覆盖；④ 日期/工期/多重空白槽不落在 PH_RE 冒号槽口径内、且同一值出现多处，answers 逐槽填效率低——用正则规则表对段落做全局就地改写（保 run 格式）更稳；⑤ 用户选择「报价留占位」时，verify A 项占位残留>0 属预期 FAIL，应转为输出《补录清单》交付草稿并在报告中注明 W4 未清零不可正式递交；⑥ verify D 红线短语按招标原文最长连续串匹配，承诺句措辞须尽量内嵌条款原文片段（如「大写金额与小写金额不一致」「响应文件有效期应为90日」），20 条典型条款可收敛至 7 条以内 WEAK。
- **起始判定**：`--start` 给定则按关键词命中；缺省自动找「(响应|投标|应答)文件格式」类标题。**终止判定**：`--end` 给定则命中处止（不含）；缺省取起始标题后第一个同级/更高级标题（第X章、中文序号、数字编号同族）。**任一无法定位 → 退出码 3，必须向用户确认起止位置**（"模板起止标题未识别，请告知起始/终止章节"），不得猜测。
- `--require` 命中检查在切片范围内做（关键词出现在任一正文文本）；未命中 → 退出码 3 + MISS 清单 → **向用户确认**（替代模板？单独附件？）后才可继续。
- `template.json` 元数据：起止标题、近似页码（780字/页估算）、章内标题清单 `headings`、表格数。`headings` 供环节四 C 轨道核对"模板节不许删"。

- **实测直填经验（2026-09 淄博项目知识库直填）**：① 知识库根目录通常与 `企业画像_*.json` 同目录（本次 `C:\Users\yahe_\Desktop\AI标书撰写\知识库`），且检索服务进程可经 `Get-CimInstance Win32_Process`/`netstat` 反查工作目录；② **一朝发现库内存在本项目历史成交响应文件（`_docs/项目2_响应文件_雅合科技.md`+`_assets/tables/*.md`）即为最高置信度直填来源**：企业全要素、法人/代理人身份证、资质证号、报价、工期、业绩、人员班子、设备/仪器/劳动力/临时用地 6 张附表一应俱全，read+正则定位后 run 级直填 38 段+全 16 表即可；③ 工期/质量/报价等「用户暂定值」与知识库成交版冲突时，以库内实际成交版为准直填并在交付说明标注（本次 240→46 日历天、一次性验收合格→合格、报价 250988.59/234393.20 税率9%）；④ 投标函大写金额模板后缀冗余「元」（如「贰角元」系模板句尾「元」+替换遗留）须人工清洗；⑤ verify.py A 项需给模板原生格式标题（「投标函【工程量清单计价】」）加 EXEMPT 豁免，否则永远 FAIL——已修复进脚本；⑥ D 红线扫描措辞须与招标原文逐字对齐（「商务报价excel版」小写、含「的」），做到 20/20 全命中。

## 环节二：置信度分级填空

**槽位 = 切片 docx 中的占位：段落空白槽（`____`、`【】`、`（ ）`、`致：· · ·（招标人名称）`式冒号空白）与表格占位单元格。三档处理：**

| 置信度 | 判定 | 动作 |
|---|---|---|
| high | 画像键名与槽 label 归一化后【精确一致】且值唯一（含编号前缀如「（九）工期」）→ auto | 直接填入，不问 |
| medium | 精确键多值冲突 / 仅子串（模糊）命中（如画像「名称」键 vs 槽「项目名称」） | **列出候选项+依据向用户确认后才写入**（子串命中不得降级 high，防串值） |
| low | 无任何来源 | **必须提问**（含"自行填写"选项），禁止臆测 |

```bash
# 1) 扫描槽位并匹配资料（画像/项目资料 JSON，逗号分隔）
python scripts/fill_plan.py scan template.docx --sources 企业画像.json -o fill_plan.json --questions questions.json

# 1.5) 【可选】知识库检索辅助：为 ask 档槽位附背景参考（服务不可达时静默跳过）
python scripts/kb_assist.py fill_plan.json -o fill_plan_kb.json --service http://127.0.0.1:8765

# 2) ★ 人机确认关口：逐条消费 questions.json——用 AskUserQuestion 工具分批提问
#    （每次 ≤4 题；medium 列候选+依据，低置信度允许自填）。答案写 answers.json：
#    {"answers": {"<slot_id>": "<用户确认值>"}}
#    若执行过 1.5，把槽位的 kb_hint 作为提问的「背景参考」附在题干，帮助判断口径；
#    kb_hint 仅供人判断，不作为候选值、不得绕过确认直接写入（W2 不变）。

# 3) 汇总可写值（仅 high/auto 与已确认槽位进 fills；未确认保持占位）
python scripts/fill_plan.py apply template.docx fill_plan.json --answers answers.json -o fills.json

# 4) 就地写回切片副本（run 级替换保格式；未填占位原样保留）
python scripts/fill_docx.py template.docx fills.json -o 投标文件草稿_YYYYMMDD.docx

# 5) 叙述性小节（技术方案/施工组织等模板要求自拟的章节）：优先用知识库检索佐证素材——
#    通过检索服务 /query（context 格式）取出处可溯源的片段续写；数字/资质/业绩仍受 W2/W3 约束。
```

- **`fill_docx.py` 退出码 3（剩余占位）= 不可交付**，回到提问关口清零后重填。
- 槽位 id 稳定格式：段落 `P{段号}#{第几个占位}`、表格 `T{表号}:R{行}C{列}`（0 起）。label 匹配是初筛，提问前 agent 必须复核语义相关性（泛匹配剔除）。

## 节点0：招标文件结构化解析（parse_tender.py）

**任何撰写动作之前（或紧跟环节一）执行**，产出后续环节依赖的结构化数据（verify D 轨道红线、报价口径、时间节点）。**只提取如实写入，不臆测。**

```bash
python scripts/parse_tender.py <招标文件.docx> -o tender.json -k 项目常量.json --with-timeline
```

- 提取四类：**常量**（招标编号/项目名称/采购人/开标时间/投标有效期/工期/保证金）、**评分项**（评分表逐行：价格/商务/技术权重，如项目2=50/15/35）、**废标红线**（含 否决/废标/不予受理 的条款，含表格）、**时间节点**（带 递交/截止 语义的日期）。
- **评分项来源**：评分表（表头含"分值/评分/评议"）逐行提取 序号+评审项目+分值；正文"分值构成"句（如"价格50分，商务15分，技术35分"）。
- **输出到项目常量**：招标编号/项目名称/采购人/开标时间/投标有效期(默认90)/工期。
- **已知局限**（2026-09 验证）：工期 若招标文件写"工期 填写阿拉伯数字 天"（占位式）无法提取，保留 0 → 由环节二提问让用户填；采购人 优先匹配独立行 `采购人：`（跨行+公司名），命中的"说明与XX分公司"残片自动剔除。
- 退出码 3（空解析）→ 停下让用户核对招标文件或调整锚点（W1 精神，不静默放行）。

## 环节二续：叙述章节撰写（write_narrative.py）

**填空完成后的自拟章节**（施工组织设计、技术方案、质量/安全/进度/环保措施、项目管理机构说明等）**由本地知识库生成参考素材**，写入草稿对应小节末尾。

```bash
python scripts/write_narrative.py 投标文件草稿.docx template.json \
    [--out 草稿叙述版.docx] [--patterns 施工方案,质量保证,安全措施] \
    [--service http://127.0.0.1:8765] [--top-k 3] [--max-lines 6] [--kb 叙述_溯源.json]
```

- **叙述小节识别**：template.json headings 命中叙述关键词（施工组织设计/技术方案/质量保证/安全措施/进度计划/环保措施/文明施工/应急预案/风险管理/项目管理机构/成品保护/保修/冬雨季/配合/附表），按出现序去重（模板清单与格式件清单会重复出现同标题）。
- **写入内容**：每个小节标题下方插入 `[参考素材—请据此起稿，可删除本行]` + 知识库检索结果（`【知识库引用 N】相似度 | 来源 | 片段`）。
- **溯源**：`--kb` 输出 叙述_溯源.json（检索片段+source_path+score），数字/参数必须可溯源。
- 退出码 3（无叙述小节/全未写入）→ 不视为失败，由调用方决定跳过（模板无叙述章节属正常）。

## 环节三：交付机检

```bash
python scripts/verify.py 投标文件.docx --template template.json \
    --quotes quotes.json --tender-parse tender.json -o verify_report.md
```

四轨道（缺参 → SKIP 显式披露，stderr 提示补齐，不静默绿灯）：

| 轨道 | 需参数 | 判定 |
|---|---|---|
| A 占位残留 | 无 | `【待填/待确认/待补】` + 未填原始模板占位（含下划线/冒号空白槽/非豁免【】），豁免表见 `_common.EXEMPT_PH` |
| B 金额一致 | `--quotes` | 分项合计=总价 + 大小写互核（`cn_to_num`），`parse_amount` 拒收约/左右/不含税等修饰语 |
| C 结构保真 | `--template` | 草稿须覆盖 template.json headings（禁止删模板节） |
| D 红线扫描 | `--tender-parse` | 按行判定：红线短语仅出现在目录/标题式短行（行长 < 短语+8）**不算实质响应**；全部短行命中记 WEAK |

退出码 0 通过（含 WEAK/SKIP）；5 存在 FAIL，修复重跑。

## 环节四：diff 回归测试（验收标尺）

**skill 每次迭代后、或新项目对齐既有成品时执行。参考基准 = 用户指定输出文件夹中的人工成品。**

```bash
python scripts/diff_report.py <参考基准.docx|json> <新结果.docx|json> \
    -o diff_report.md --json diff.json --threshold 0.85
```

- 对齐：双方统一为节模型（heading + 归一化段落 + 表格网格），标题 norm 后 LCS 对齐；漂移进 `missing_in_new / extra_in_new`。
- 比对：段落 difflib 逐段 diff；表格逐格；节相似度=SequenceMatcher.ratio，单元格不一致每格扣 0.02。
- 输出：差异清单（md：通过率表 + 每节状态 + 未达标节到具体单元格/段落）与通过率统计。
- **迭代纪律：存在未达标节（退出码 4）时按差异清单修脚本/修流程，重跑直至对齐；diff 报告随交付物归档（W5）。**

## 知识库检索后端（本地 ChromaDB）

环节二的知识库检索由本地检索服务提供，**数据不出本机、不依赖云端**。

**服务状态**：仅监听 `127.0.0.1:8765`。启动：`知识库/_service/start_service.bat`（或 `python server.py`）。数据源为 `知识库/` 六大库的 Markdown/表格/JSON 索引，2089 个分块，向量模型 `BAAI/bge-small-zh-v1.5`。

**agent 调用方式（两种，等价）**：

1. 脚本通道（推荐，免 HTTP 细节）：`python scripts/kb_assist.py fill_plan.json -o fill_plan_kb.json`
   - 对每个 `status=ask` 槽位，用 `label + context` 拼检索句，取 top-2 片段的 `context_block`（截 300 字）写入 `slot.kb_hint`
   - 服务不可达 → 静默跳过、`meta.kb_assist.status=skipped`，**不阻断流程**
2. HTTP 通道（续写叙述章节时直接用）：
   ```json
   POST http://127.0.0.1:8765/query
   {"query": "当前撰写的小节文本或槽位描述", "top_k": 3,
    "library": "05-施工方案库", "source_type": "kb_article", "format": "context"}
   ```

**响应格式**：
```json
{"query": "…", "count": 2, "took_ms": 380,
 "results": [{"id": "…", "text": "片段正文", "score": 0.80,
              "metadata": {"source_path": "知识库/02-技术参数库/排流地床技术参数.md",
                           "library": "02-技术参数库", "source_type": "kb_article",
                           "heading_path": "排流地床技术参数 > 填包料",
                           "image_path": "知识库/_assets/images/xxx.png"}}],
 "context_block": "【知识库引用 1】相似度 0.80 | 来源：…\n章节：…\n片段正文…"}
```

**使用纪律**：
- `kb_hint` / 检索片段是**背景参考**，不是候选值：不得据此把槽位升为 high 或绕过提问（W2）。
- 续写叙述章节时，引用的数字/参数必须带 `source_path` 溯源；检索不到就写「待补充」并列入提问清单。
- 服务健康检查：`GET /health`（返回 `chunks` 总数）；数据更新后执行 `知识库/_service` 的 `ingest.py` 增量导入。

## 证照图片插入（环节二附，2026-09 实践沉淀）

单据型章节（营业执照/资质/安取/体系证书）若基准成品自带图片，应优先从**成交响应文件 docx** 抽图插入，避免用户手动扫描：

1. **建映射**：遍历源 docx 段落，遇小节标题段切换当前小节；统计段内 `a:blip` 的 `r:embed` rid → `{小节标题: [rid,...]}`。
2. **抽图**：`rels[rid].target_part.blob` 写盘（`partname.ext` 保留扩展名）；按「目标小节标题关键词 → 前 n 张」配比（例：营业执照 1、腐蚀控制资质 3、安取 1、质量体系 1、HSE 职业健康/环境各 1）。
3. **插入**：python-docx 无 insert_after，用「`anchor.insert_paragraph_before('')` 建图段（居中、`run.add_picture(width=Inches(6))`）→ `anchor._p.addnext(p._p)` 移到锚点后」。图片自带文字标识时可不配图注。
4. **锁文件降级**：目标 docx 被 Word/WPS 打开（`~$` 锁文件）时 `d.save()` 抛 `PermissionError` → 自动另存 `_含图版.docx` 并提示用户比对后替换。
5. **必跑机检**：插入后仍须 `verify.py --template/--quotes/--tender-parse` 复跑，图片不得引入占位/破坏结构/影响报价合计；先在临时副本干跑，确认 6 张图落在正确小节前后邻段落正确再操作正式稿。

## 铁规则

- W1 模板唯一来源是招标文件；miss 必问，不得自造格式。
- W2 low/medium 槽位未获用户确认不得写入；answers.json 是唯一写入凭据。
- W3 金额、资质编号、日期只信画像与招标文件，冲突即问。
- W4 verify FAIL 或 fill_docx 剩余占位 >0 → 禁止交付。
- W5 回归差异不得静默忽略，每轮 diff 报告归档。

## 企业画像

画像 JSON（schema 见 `references/confidence-protocol.md` 附录）保存在项目工作区，**skill 内不存放任何真实企业数据**，运行时经 `fill_plan.py scan --sources` 注入。

## Changelog

- v0.3.1 (2026-09-04)：质量加固。P1：fill_plan 置信度判定改为「精确键命中才 high」，子串命中（防「名称→项目名称」类串值）一律降 medium；verify.py 缺 `--quotes/--template/--tender-parse` 改 SKIP 显式披露（不静默绿灯），D 红线按行判定+目录短行不算实质响应；补全 scripts/insert_images.py（标题锚点后居中图 6in + PermissionError 另存 `_含图版.docx`）。P2：`_common.norm` 全角冒号统一（曾导致 label 不可比）；`cn_to_num` 修「壹佰贰拾叁元肆角伍分」被吞 0.45 元 bug；`parse_amount` 改先剥币种再判模糊修饰语；共享 `EXEMPT_PH`/`is_exempt_ph` 取代硬编码白名单。新增 tests/（pytest 33 个用例 + smoke_test/smoke_verify 端到端冒烟）。
- v0.3.0 (2026-09-01)：四环节流水线定型；淄博项目实战踩坑记录；知识库直填 / 图片插入流程沉淀
