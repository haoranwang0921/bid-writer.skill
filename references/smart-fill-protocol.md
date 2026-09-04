# 模型检索填空协议

本协议用于环节二的智能填空。目标是让运行本 Skill 的模型理解槽位语义、主动查询数据库，再把有证据的唯一事实交给确定性脚本写回。模型负责理解与检索，`smart_fill.py validate` 负责放行，`fill_docx.py` 负责保格式写回。

## 数据源目录

可选的 `data_sources.json` 告诉模型有哪些只读数据源。`fill_plan.py scan --sources` 使用的 JSON 会自动登记进 `fill_plan.meta.source_catalog`；SQLite、HTTP 检索服务或其他连接器可手工补充：

```json
{
  "sources": [
    {
      "id": "company-profile",
      "kind": "json",
      "location": "C:/project/企业画像.json",
      "source_type": "structured",
      "access": "read_only"
    },
    {
      "id": "tender-facts",
      "kind": "json",
      "location": "C:/project/tender.json",
      "source_type": "tender",
      "access": "read_only"
    },
    {
      "id": "local-kb",
      "kind": "http_search",
      "location": "http://127.0.0.1:8765/query",
      "source_type": "knowledge_base",
      "access": "read_only"
    }
  ]
}
```

`source_type` 取值：

| 类型 | 含义 | 可自动填入 |
|---|---|---|
| `structured` | 企业画像、业务数据库等结构化事实 | 是，须唯一且可定位 |
| `tender` | 当前招标文件的结构化解析结果 | 是，须唯一且可定位 |
| `confirmed` | 用户已确认并固化的项目事实 | 是，须唯一且可定位 |
| `document` | 原始文档页、节、段中的事实 | 否，先让用户确认 |
| `knowledge_base` | 向量检索或语义检索片段 | 否，只作候选与背景 |

数据库访问必须只读。模型不得修改企业库、项目库或知识库来制造可填写事实。

## 执行流程

### 1. 生成模型任务

```bash
python scripts/smart_fill.py prepare fill_plan.json \
  --catalog data_sources.json -o model_tasks.json
```

缺少额外目录时省略 `--catalog`，脚本会使用 `fill_plan.meta.source_catalog`。默认只生成 `status=ask` 的任务；`--include-auto` 可把原机械 high 槽位一并交给模型审计，但模型结果不会覆盖原机械 high。

### 2. 模型检索

模型逐个读取 `model_tasks.tasks`，结合 `label + context + safety_class` 判断字段语义，并使用当前环境可用的只读工具查询数据源：

1. 优先查当前项目的 `tender` 与 `confirmed` 数据。
2. 企业固有事实查 `structured` 数据。
3. 结构化数据没有结果时再查 `document` 与 `knowledge_base`。
4. 查询结果冲突时全部保留，不得自行选择较顺眼的值。
5. 没找到事实时返回 `low`，不得生成补全值。

金额、报价、日期、工期、证照编号、统一社会信用代码、银行账号、法定代表人等 `critical_fact` 必须核对当前项目口径、有效期和适用主体。历史文件或向量知识库命中不能直接作为自动填写依据。

### 3. 提交模型决策

模型把每个槽位的查询过程写入 `model_decisions.json`：

```json
{
  "decisions": [
    {
      "slot_id": "P12#0",
      "canonical_field": "basic.legal_representative",
      "query": "在当前投标主体企业画像中查询法定代表人",
      "selected_value": "张三",
      "confidence": "high",
      "reason": "当前投标主体只有一个有效法定代表人记录",
      "evidence": [
        {
          "value": "张三",
          "source": "company-profile",
          "source_type": "structured",
          "locator": "companies/9137.../basic/legal_representative",
          "record_id": "company-9137...",
          "updated_at": "2026-08-20"
        }
      ]
    }
  ]
}
```

`locator` 必须让下一位审核者能回到原记录，不能只写“数据库”“知识库”“搜索结果”等笼统描述。`selected_value` 必须逐字出现在某条 evidence 的 `value` 中。

### 4. 确定性验证

```bash
python scripts/smart_fill.py validate fill_plan.json model_decisions.json \
  --catalog data_sources.json -o smart_fill.json --questions smart_questions.json
```

如果 `prepare` 时没有传额外目录，这里也省略 `--catalog`，两步都会使用 `fill_plan.meta.source_catalog`。若 `prepare` 使用了外部目录，`validate` 必须传入同一文件；未登记来源或来源类型与目录不一致的证据只能进入待确认候选，不能自动填写。

只有同时满足以下条件的模型 high 决策进入 `auto_fills`：

- 槽位 id 存在且无重复决策；
- `selected_value` 与证据值完全一致；
- 所有有效证据只有一个不同值；
- 至少一条选中值证据来自 `structured`、`tender` 或 `confirmed`；
- 证据的 `source` 已登记在数据源目录中，且声明类型与目录一致；
- 证据包含 `source` 和可复查的 `locator`；
- 模型记录了实际查询 `query` 和选择理由 `reason`。

证据冲突、仅知识库命中、只有历史文档、缺定位、值不在证据中等情况自动降为 medium/low，并进入 `smart_questions.json`，不会进入 Word。

### 5. 合并并写回

```bash
python scripts/fill_plan.py apply template.docx fill_plan.json \
  --smart smart_fill.json --answers answers.json -o fills.json
python scripts/fill_docx.py template.docx fills.json -o 投标文件草稿.docx
```

写入优先级：

1. `answers.json` 中的用户确认值；
2. `smart_fill.json` 中经验证的模型证据值；
3. `fill_plan.py scan` 原有的机械精确匹配 high 值；
4. 其余保持占位。

`fills.json.audit` 为每个值记录 `human_answer`、`model_evidence` 或 `deterministic_auto` 来源。模型证据的 query、reason、canonical_field 和 evidence 一并归档，供交付审计。

## 边界

- 模型可以决定“去哪里查、怎样组合查询、哪个字段与槽位语义对应”，不能把常识或语言概率当作企业事实。
- 向量相似度代表文本相关性，不代表事实正确性或当前有效性。
- 用户回答可以覆盖模型与机械匹配；覆盖记录必须留在 `answers.json`。
- `smart_fill.py` 不直接调用云端模型，不保存 API Key。运行 Skill 的 agent 即模型层，因此该协议可适配 Codex、WorkBuddy 或其他支持工具调用的 agent。
- 智能填空通过后仍必须运行 `fill_docx.py` 残留检查和 `verify.py` 四轨机检。
