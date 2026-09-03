#!/usr/bin/env python3
"""diff_report.py — 逐节 diff 回归对比（环节三）

以既有参考输出为基准，对新生成结果逐节 diff，输出差异清单（字段级不一致、
缺失/多余章节）与通过率统计，用于迭代修正直至对齐。

支持三种对比模式：
  1) json↔json ：两份投标内容/模板 JSON（document.sections[].heading + 归一化文本/表格网格）
  2) docx↔docx：两份 Word 成稿（提取标题层级 + 段落 + 表格后逐节对齐）
  3) json↔docx：新结果 JSON 对比参考 docx（自动识别扩展名）

CLI:
  python diff_report.py <参考基准> <新结果> -o diff_report.md [--json diff.json] \
      [--ignore 页眉,页脚,目录] [--threshold 0.85]

对齐策略：按章节标题 norm 后做最长公共子序列匹配；标题漂移的节进入
「missing_in_new / extra_in_new」。每节按字段级比对：
  - heading 一致
  - 段落：归一化后 difflib 逐段 diff，统计 equal/replace/insert/delete
  - 表格：逐格 grid 比对，统计不一致单元格
节内相似度 = 2*matched/(len_a+len_b)（SequenceMatcher.ratio）。
通过率 = 相似度>=threshold 且无字段级 FAIL 的节数 / 参考总节数。
退出码：0 全部达标；4 存在低于 threshold 的节（供流水线拦截）；2 输入错误。
"""
import argparse, difflib, json, re, sys
from datetime import datetime
sys.path.insert(0, __import__("os").path.dirname(__file__))
from _common import atomic_write_json, guard_out, norm

try:
    from docx import Document
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxPara
except ImportError:
    print("ERROR: python-docx 未安装", file=sys.stderr)
    sys.exit(2)

HEADING_STYLE_RE = re.compile(r"(?:Heading|标题)\s*([1-9])", re.I)
CN = "零一二三四五六七八九十百千"
CHAPTER_RE = re.compile(r"^第\s*[" + CN + r"0-9]+\s*[章节部分篇卷]")
NUM_RE = re.compile(r"^(?:\d{1,2}(?:[.．]\d{1,2}){0,3})[、.．\s]\s*\S")


def docx_to_model(path):
    """docx → {sections:[{heading, level, paragraphs:[...], tables:[[...]]}]}"""
    doc = Document(path)
    sections, cur = [], None
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            p = DocxPara(child, doc)
            text = p.text.strip()
            if not text:
                continue
            lv = None
            m = HEADING_STYLE_RE.search(p.style.name or "")
            if m:
                lv = int(m.group(1))
            elif CHAPTER_RE.match(text) or (NUM_RE.match(text) and len(text) <= 40):
                lv = 2 if re.match(r"^\d+\.\d", text) else 1
            if lv:
                cur = {"heading": text, "level": lv, "paragraphs": [], "tables": []}
                sections.append(cur)
            else:
                if cur is None:
                    cur = {"heading": "(文首)", "level": 0, "paragraphs": [], "tables": []}
                    sections.append(cur)
                cur["paragraphs"].append(text)
        elif child.tag.endswith("}tbl"):
            t = DocxTable(child, doc)
            if cur is None:
                cur = {"heading": "(文首)", "level": 0, "paragraphs": [], "tables": []}
                sections.append(cur)
            cur["tables"].append([[c.text.strip() for c in r.cells] for r in t.rows])
    return {"format": "docx", "sections": sections}


def json_to_model(path):
    """投标内容.json / 模板.json / 填充.json → 统一模型。"""
    d = json.load(open(path, encoding="utf-8"))
    secs = []
    # 兼容 bid-studio 投标内容.json（document.sections 带 paragraphs 字符串数组）
    src = None
    if isinstance(d, dict) and "document" in d and "sections" in d["document"]:
        src = d["document"]["sections"]
    elif isinstance(d, dict) and "sections" in d:
        src = d["sections"]
    if not src:
        raise ValueError("无法识别 JSON 结构（缺 document.sections 或 sections）")
    for s in src:
        heading = s.get("heading") or s.get("title") or ""
        paras = []
        for p in s.get("paragraphs", []):
            if isinstance(p, str):
                paras.append(p)
            else:
                paras.append(p.get("text", ""))
        tables = []
        for t in s.get("tables", []):
            tables.append(t.get("grid") or [])
        secs.append({"heading": heading, "level": int(s.get("level") or 1) if str(s.get("level", "1")).isdigit() else 1,
                     "paragraphs": paras, "tables": tables})
    return {"format": "json", "sections": secs}


def load_model(path):
    if path.lower().endswith(".docx"):
        return docx_to_model(path)
    return json_to_model(path)


def _sec_signature(sec):
    return norm(sec["heading"])


def align(ref_secs, new_secs, ignore):
    """按标题 LCS 对齐节；返回 [(ref_sec|None, new_sec|None)]。"""
    rkeys = [_sec_signature(s) for s in ref_secs]
    nkeys = [_sec_signature(s) for s in new_secs]
    sm = difflib.SequenceMatcher(a=rkeys, b=nkeys, autojunk=False)
    pairs = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                pairs.append((ref_secs[i1 + k], new_secs[j1 + k]))
        elif tag == "replace":
            rl, nl = list(range(i1, i2)), list(range(j1, j2))
            for k in range(max(len(rl), len(nl))):
                r = ref_secs[rl[k]] if k < len(rl) else None
                n = new_secs[nl[k]] if k < len(nl) else None
                pairs.append((r, n))
        elif tag == "delete":
            for i in range(i1, i2):
                pairs.append((ref_secs[i], None))
        elif tag == "insert":
            for j in range(j1, j2):
                pairs.append((None, new_secs[j]))
    return pairs


def compare_section(ref, new, ignore, threshold):
    """字段级比对：返回差异记录。"""
    rec = {"heading": (ref or new)["heading"],
           "status": "ok", "similarity": 0.0, "diffs": []}
    if ref and not new:
        rec["status"] = "missing_in_new"; rec["similarity"] = 0.0
        rec["diffs"].append({"type": "section_missing", "detail": "参考有、新结果缺该节"})
        return rec
    if new and not ref:
        rec["status"] = "extra_in_new"; rec["similarity"] = 0.0
        rec["diffs"].append({"type": "section_extra", "detail": "新结果多出的节（可能模板外臆造）"})
        return rec
    rp = [norm(x) for x in ref["paragraphs"]]
    np_ = [norm(x) for x in new["paragraphs"]]
    sm = difflib.SequenceMatcher(a=rp, b=np_, autojunk=False)
    ratio = sm.ratio()
    # 空段落节（纯标题/目录）以标题匹配为准
    if not rp and not np_:
        ratio = 1.0 if norm(ref["heading"]) == norm(new["heading"]) else 0.0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ("insert", "delete", "replace"):
            for i in range(i1, i2):
                if tag in ("delete", "replace"):
                    rec["diffs"].append({"type": "para_" + tag, "side": "ref", "text": ref["paragraphs"][i][:120]})
            for j in range(j1, j2):
                if tag in ("insert", "replace"):
                    rec["diffs"].append({"type": "para_" + tag, "side": "new", "text": new["paragraphs"][j][:120]})
    # 表格逐格
    rtabs, ntabs = ref["tables"], new["tables"]
    if len(rtabs) != len(ntabs):
        rec["diffs"].append({"type": "table_count", "detail": f"参考 {len(rtabs)} 表 / 新 {len(ntabs)} 表"})
    for ti in range(min(len(rtabs), len(ntabs))):
        rt, nt = rtabs[ti], ntabs[ti]
        rows = max(len(rt), len(nt))
        for r in range(rows):
            rrow = rt[r] if r < len(rt) else []
            nrow = nt[r] if r < len(nt) else []
            cols = max(len(rrow), len(nrow))
            for c in range(cols):
                rv = norm(rrow[c]) if c < len(rrow) else "<无>"
                nv = norm(nrow[c]) if c < len(nrow) else "<无>"
                if rv != nv:
                    rec["diffs"].append({"type": "cell", "table": ti, "cell": f"R{r}C{c}",
                                         "ref": (rrow[c] if c < len(rrow) else "")[:80],
                                         "new": (nrow[c] if c < len(nrow) else "")[:80]})
                    ratio = max(0.0, ratio - 0.02)
    rec["similarity"] = round(ratio, 3)
    hard = [d for d in rec["diffs"] if d["type"] in ("section_missing", "section_extra", "cell", "table_count", "para_delete", "para_insert")]
    if ratio < threshold or hard:
        rec["status"] = "fail"
    return rec


def render_md(report, out_md):
    s = report["summary"]
    lines = [
        f"# 标书回归 diff 报告", "",
        f"- 参考基准：`{report['ref']}`（{s['ref_sections']} 节）",
        f"- 新结果：`{report['new']}`（{s['new_sections']} 节）",
        f"- 生成时间：{report['generated_at']}",
        f"- 阈值：相似度 ≥ {report['threshold']} 且无字段级 FAIL 视为对齐", "",
        "## 通过率统计", "",
        f"| 指标 | 值 |", f"|---|---|",
        f"| 对齐节数 / 参考总节数 | {s['passed']} / {s['ref_sections']} |",
        f"| **通过率** | **{s['pass_rate']:.1%}** |",
        f"| 缺失节 (missing_in_new) | {s['missing']} |",
        f"| 多余节 (extra_in_new) | {s['extra']} |",
        f"| 字段级差异条数 | {s['total_diffs']} |",
        f"| 结论 | {'✅ 全部达标' if s['all_pass'] else '❌ 存在未对齐节，需迭代'} |", "",
        "## 差异清单", "",
        "| # | 章节 | 状态 | 相似度 | 差异要点 |", "|---|---|---|---|---|",
    ]
    for i, r in enumerate(report["sections"], 1):
        summ = {}
        for d in r["diffs"]:
            summ[d["type"]] = summ.get(d["type"], 0) + 1
        desc = "，".join(f"{k}×{v}" for k, v in summ.items()) or "—"
        head = r["heading"][:28]
        lines.append(f"| {i} | {head} | {r['status']} | {r['similarity']:.2f} | {desc} |")
    fail_secs = [r for r in report["sections"] if r["status"] == "fail"]
    if fail_secs:
        lines += ["", "## 未达标节明细", ""]
        for r in fail_secs:
            lines.append(f"### 「{r['heading']}」（相似度 {r['similarity']:.2f}）")
            for d in r["diffs"][:40]:
                if d["type"] == "cell":
                    lines.append(f"- 单元格 `{d['cell']}`：参考=`{d['ref']}` → 新=`{d['new']}`")
                elif d["type"].startswith("para_"):
                    lines.append(f"- 段落[{d['side']}] {d['text'][:80]}")
                else:
                    lines.append(f"- {d['type']}：{d.get('detail','')}")
            lines.append("")
    with open(out_md, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ref", help="参考基准（既有输出 docx/json）")
    ap.add_argument("new", help="新结果（docx/json）")
    ap.add_argument("-o", "--out", default="diff_report.md")
    ap.add_argument("--json", dest="jsonout", default="")
    ap.add_argument("--ignore", default="页眉,页脚,目录")
    ap.add_argument("--threshold", type=float, default=0.85)
    a = ap.parse_args()
    ignore = [norm(x) for x in a.ignore.split(",") if x.strip()]
    try:
        rmod = load_model(a.ref)
        nmod = load_model(a.new)
    except Exception as e:
        print(f"ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(2)

    def keep(sections):
        return [s for s in sections if not any(ig in norm(s["heading"]) for ig in ignore)]

    ref_secs, new_secs = keep(rmod["sections"]), keep(nmod["sections"])
    pairs = align(ref_secs, new_secs, ignore)
    results = [compare_section(r, n, ignore, a.threshold) for r, n in pairs]
    passed = sum(1 for r in results if r["status"] == "ok")
    missing = sum(1 for r in results if r["status"] == "missing_in_new")
    extra = sum(1 for r in results if r["status"] == "extra_in_new")
    total_diffs = sum(len(r["diffs"]) for r in results)
    report = {
        "ref": a.ref, "new": a.new,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "threshold": a.threshold,
        "summary": {"ref_sections": len(ref_secs), "new_sections": len(new_secs),
                    "passed": passed, "pass_rate": passed / max(len(ref_secs), 1),
                    "missing": missing, "extra": extra, "total_diffs": total_diffs,
                    "all_pass": passed == len(ref_secs) and extra == 0},
        "sections": results,
    }
    out = guard_out(a.out)
    render_md(report, out)
    if a.jsonout:
        atomic_write_json(guard_out(a.jsonout), report)
    print(f"OK pass_rate={report['summary']['pass_rate']:.1%} passed={passed}/{len(ref_secs)} "
          f"missing={missing} extra={extra} diffs={total_diffs} -> {out}")
    if not report["summary"]["all_pass"]:
        sys.exit(4)


if __name__ == "__main__":
    main()
