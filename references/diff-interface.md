# 环节三/四：diff 回归对比实现逻辑与接口

## 统一节模型

无论输入是 docx（人工成品基准）还是 JSON（机器产物），都先归一为：

```jsonc
{"format":"docx|json", "sections":[{"heading","level","paragraphs":[…],"tables":[[[…]]]}]}
```

- docx：`docx_to_model` 按样式/文本启发式重建节（与 extract_template 同一标题判定逻辑，保证两侧口径一致）；
- json：兼容 `document.sections[]`（bid-studio 风格）与 `sections[]`（template/filled 风格）；段落兼容 str 与 `{text}`。

## 对齐与比对

1. **节对齐**：双方标题 norm 后 `difflib.SequenceMatcher` 取等值段；replace 段内按位置配对；未配对的进 `missing_in_new` / `extra_in_new`（extra=模板外臆造嫌疑，直接 FAIL）。
2. **段比对**：节内段落 norm 序列逐段 diff，记 `para_replace/insert/delete`（ref 侧文本入差异清单）。
3. **表比对**：网格逐格 norm 相等校验，不一致记 `{ref, new}` 单元格明细；表数量不等记 `table_count`。每个 cell 差异另扣节相似度 0.02。
4. **节相似度**：段落 SequenceMatcher.ratio；双方无段落时按标题一致性给 1/0。
5. **忽略项**：`--ignore 页眉,页脚,目录`（默认）——标题含关键词的节不参与统计（机器成稿不渲染目录是预期行为）。

## 通过判定

- 节级 PASS：`similarity ≥ threshold（默认 0.85）` 且无硬差异（缺失/多余/单元格/段落增删）；
- **总通过率 = PASS 节数 / 参考节数**；`all_pass` 要求参考全 PASS 且零多余节；
- 退出码：0 全达标；4 有未达标节（迭代信号）；2 输入错误。

## 迭代工作流

```
diff_report.py → 差异清单
  ├─ 结构性缺失（missing_in_new）→ 修 extract/keep 范围或 render 丢节
  ├─ 字段级不一致（cell/para_replace）→ 修画像数据 / fill 匹配逻辑 / 叙述生成措辞
  ├─ extra_in_new → 检查环节二是否写入了无来源内容（W2 违例）
  └─ 通过率 <100% → 修正后重跑整条流水线（模板→填空→渲染→diff）
```

报告（md）必含：通过率统计表、每节状态一览、未达标节明细（到单元格与段落原文）。差异不可静默忽略（W5）——确属参考版本身过时/格式漂移的节，在报告外人工批注后方可调阈值，并注明理由。
