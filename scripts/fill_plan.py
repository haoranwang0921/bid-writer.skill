#!/usr/bin/env python3
"""fill_plan.py v2 — 置信度分级填空（对切片 template.docx 操作）

槽位 = template.docx 中的占位（段落空白槽/下划线/【】 与表格占位单元格）。
CLI:
  python fill_plan.py scan <template.docx> [--sources a.json,b.json] \
      -o fill_plan.json [--questions questions.json]
  python fill_plan.py apply <template.docx> <fill_plan.json> [--answers answers.json] -o fills.json

置信度：high=来源唯一命中→auto；medium=2-4候选→ask；low=无来源→ask（禁止臆测）。
fills.json：{"fills": {"P12#0": "值", "T3:R1C2": "值"}}（仅 high/已确认槽位）。
"""
import argparse, json, re, sys, os
from datetime import datetime
sys.path.insert(0, os.path.dirname(__file__))
from _common import atomic_write_json, guard_out, norm

try:
    from docx import Document
except ImportError:
    print("ERROR: python-docx 未安装", file=sys.stderr)
    sys.exit(2)

PH_RE = re.compile(r"_{2,}|＿{2,}|×{3,}|【[^】]{0,20}】|（\s*）|\(\s*\)|[：:][ \t\u3000]{2,}(?=[（(年元天日%]|$)")
UNFILLABLE_RE = re.compile(r"待确认|待补|见扫描件|按招标格式")  # 画像中的占位值不得视为高置信度


def flatten(obj, prefix=""):
    out = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            key = f"{prefix}.{k}" if prefix else str(k)
            out.update(flatten(v, key) if isinstance(v, (dict, list)) else {key: v})
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            out.update(flatten(v, f"{prefix}[{i}]"))
    else:
        out[prefix or "value"] = obj
    return out


def build_slots(doc):
    slots = []
    for pi, p in enumerate(doc.paragraphs):
        t = p.text
        for k, m in enumerate(PH_RE.finditer(t or "")):
            lead = norm(t[max(0, m.start() - 18):m.start()])[-16:]
            slots.append({"id": f"P{pi}#{k}", "type": "paragraph",
                          "label": lead or norm(t)[:20], "current": m.group(0),
                          "context": norm(t)[:70]})
    for ti, tb in enumerate(doc.tables):
        header = [c.text.strip() for c in tb.rows[0].cells] if len(tb.rows) else []
        seen_row = set()
        for r, row in enumerate(tb.rows):
            for c, cell in enumerate(row.cells):
                if (r, id(cell._tc)) in seen_row:
                    continue
                seen_row.add((r, id(cell._tc)))
                txt = cell.text.strip()
                if txt and PH_RE.search(txt):
                    lab = norm(header[c])[:24] if c < len(header) and header[c] else norm(txt)[:24]
                    slots.append({"id": f"T{ti}:R{r}C{c}", "type": "table_cell",
                                  "label": lab, "current": txt[:40]})
    return slots


def match_sources(slots, sources):
    idx = {}
    for path in sources:
        try:
            data = json.load(open(path, encoding="utf-8"))
        except Exception as e:
            print(f"WARN 资料读取失败 {path}: {e}", file=sys.stderr)
            continue
        for k, v in flatten(data).items():
            if v is None or isinstance(v, bool) or str(v).strip() == "":
                continue
            if UNFILLABLE_RE.search(str(v)) or str(k).lower().endswith(("profile_id", "review_status")):
                continue  # 画像占位值/元字段不作为可填来源（自然落入提问档）
            idx.setdefault(norm(k), []).append({"value": str(v), "source": os.path.basename(path)})
    for s in slots:
        lab = s["label"]
        cands, seen = [], set()
        if len(lab) >= 2:
            for key, vals in idx.items():
                if key and len(key) >= 2 and (lab in key or key in lab):
                    for v in vals:
                        if v["value"] not in seen:
                            seen.add(v["value"])
                            cands.append(v)
        s["candidates"] = cands[:4]
        uniq = {c["value"] for c in cands}
        s["confidence"] = "high" if len(uniq) == 1 else ("medium" if uniq else "low")
        s["status"] = "auto" if s["confidence"] == "high" else "ask"
    return slots


def cmd_scan(a):
    doc = Document(a.template)
    sources = [p.strip() for p in a.sources.split(",") if p.strip()] if a.sources else []
    slots = match_sources(build_slots(doc), sources)
    stats = {"total": len(slots), "high": 0, "medium": 0, "low": 0}
    for s in slots:
        stats[s["confidence"]] += 1
    atomic_write_json(guard_out(a.out), {"meta": {"template": a.template,
                     "generated_at": datetime.now().isoformat(timespec="seconds"), "stats": stats},
                     "slots": slots})
    if a.questions:
        qs = [{"id": s["id"], "label": s["label"], "current": s["current"],
               "context": s.get("context", ""), "confidence": s["confidence"],
               "options": [f"{c['value']}（依据：{c['source']}）" for c in s["candidates"]]
                          + (["自行填写"] if s["confidence"] == "low" else []),
               "basis": "来源冲突需裁决" if s["confidence"] == "medium" else "无来源，禁止臆测"}
              for s in slots if s["status"] == "ask"]
        atomic_write_json(guard_out(a.questions), {"questions": qs})
    asks = sum(1 for s in slots if s["status"] == "ask")
    print(f"OK slots={stats['total']} high={stats['high']} medium={stats['medium']} "
          f"low={stats['low']} 待确认={asks} -> {a.out}")


def cmd_apply(a):
    doc = Document(a.template)
    plan = json.load(open(a.plan, encoding="utf-8"))
    answers = {}
    if a.answers:
        raw = json.load(open(a.answers, encoding="utf-8"))
        answers = raw.get("answers", raw)
    fills, unconf = {}, 0
    for s in plan["slots"]:
        if s["id"] in answers:
            fills[s["id"]] = str(answers[s["id"]])
        elif s["status"] == "auto":
            fills[s["id"]] = s["candidates"][0]["value"]
        else:
            unconf += 1
    atomic_write_json(guard_out(a.out), {"fills": fills, "unconfirmed": unconf,
                     "generated_at": datetime.now().isoformat(timespec="seconds")})
    print(f"OK fills={len(fills)} 未确认={unconf} -> {a.out}")
    if unconf:
        print("WARN 未确认槽位将保持占位原样，交付前必须清零", file=sys.stderr)


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    s1 = sub.add_parser("scan")
    s1.add_argument("template")
    s1.add_argument("--sources", default="")
    s1.add_argument("-o", "--out", default="fill_plan.json")
    s1.add_argument("--questions", default="")
    s1.set_defaults(fn=cmd_scan)
    s2 = sub.add_parser("apply")
    s2.add_argument("template")
    s2.add_argument("plan")
    s2.add_argument("--answers", default="")
    s2.add_argument("-o", "--out", default="fills.json")
    s2.set_defaults(fn=cmd_apply)
    a = ap.parse_args()
    try:
        a.fn(a)
    except Exception as e:
        print(f"ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
