#!/usr/bin/env python3
"""fill_docx.py — 将确认值就地写回模板切片（局部修改，格式继承自源文件）

读取 template.docx 副本 + fills.json，把已确认值替换进对应占位；未确认占位
保持原样。输出即为投标文件底稿（后续叙述内容由 agent 在同文件上按模板小节
续写，禁止增删模板既有节）。

CLI:
  python fill_docx.py <template.docx> <fills.json> -o 投标文件草稿_YYYYMMDD.docx
退出码：0 全部填入；3 仍有未填占位（交付前必须清零）。
"""
import argparse, json, re, sys, os
sys.path.insert(0, os.path.dirname(__file__))
from _common import guard_out

try:
    from docx import Document
except ImportError:
    print("ERROR: python-docx 未安装", file=sys.stderr)
    sys.exit(2)

PH_RE = re.compile(r"_{2,}|＿{2,}|×{3,}|【[^】]{0,20}】|（\s*）|\(\s*\)|[：:][ \t\u3000]{2,}(?=[（(年元天日%]|$)")


def replace_in_para(p, old, new):
    """run 级替换优先；跨 run 则整段重写（记格式损失）。"""
    for r in p.runs:
        if old in r.text:
            r.text = r.text.replace(old, new, 1)
            return True
    full = p.text
    if old in full:
        new_full = full.replace(old, new, 1)
        if p.runs:
            p.runs[0].text = new_full
            for r in p.runs[1:]:
                r.text = ""
            return True
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("fills")
    ap.add_argument("-o", "--out", default=None)
    a = ap.parse_args()
    doc = Document(a.template)
    fills = json.load(open(a.fills, encoding="utf-8"))
    fmap = fills.get("fills", fills) if isinstance(fills, dict) else {}
    done = miss = 0
    # 段落槽
    for pid, val in list(fmap.items()):
        m = re.match(r"^P(\d+)#(\d+)$", pid)
        if not m:
            continue
        pi, k = int(m.group(1)), int(m.group(2))
        if pi >= len(doc.paragraphs):
            miss += 1
            continue
        p = doc.paragraphs[pi]
        ms = list(PH_RE.finditer(p.text or ""))
        if k < len(ms):
            if replace_in_para(p, ms[k].group(0), str(val)):
                done += 1
            else:
                miss += 1
    # 表格槽：建立 (ti,r,c)→cell 稳定映射
    tcoord = {}
    for ti, tb in enumerate(doc.tables):
        for r, row in enumerate(tb.rows):
            for c, cell in enumerate(row.cells):
                tcoord[(ti, r, c)] = cell
    for pid, val in list(fmap.items()):
        m = re.match(r"^T(\d+):R(\d+)C(\d+)$", pid)
        if not m:
            continue
        cell = tcoord.get((int(m.group(1)), int(m.group(2)), int(m.group(3))))
        if cell is None:
            miss += 1
            continue
        for p in cell.paragraphs:
            for ph in list(PH_RE.finditer(p.text or "")):
                if replace_in_para(p, ph.group(0), str(val)):
                    done += 1
                    break
                break
    out = a.out or re.sub(r"\.docx$", "", a.template) + "_草稿.docx"
    out = guard_out(out)
    doc.save(out)
    residual = 0
    chk = Document(out)
    for p in chk.paragraphs:
        residual += len(PH_RE.findall(p.text or ""))
    for tb in chk.tables:
        for row in tb.rows:
            for cell in row.cells:
                residual += len(PH_RE.findall(cell.text or ""))
    print(f"OK -> {out} 填入={done} 定位失败={miss} 剩余占位={residual}")
    if residual:
        print("WARN 剩余占位须在提问确认后清零，否则不可交付（W2/W4）", file=sys.stderr)
        sys.exit(3)


if __name__ == "__main__":
    main()
