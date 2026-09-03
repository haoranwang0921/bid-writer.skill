#!/usr/bin/env python3
"""_common.py — bid-writer 共享工具：原子写、路径守卫、归一化、金额解析。"""
import json, os, re, tempfile

CN_DIG = {"零": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5, "陆": 6, "柒": 7, "捌": 8, "玖": 9}
CN_UNIT = {"拾": 10, "佰": 100, "仟": 1000}


def atomic_write_json(path, obj):
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=d, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=1)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def guard_out(path):
    """输出限定：工作目录 / 系统临时目录 / 用户 .workbuddy 内，防越界写。"""
    ap = os.path.abspath(path)
    allowed = [os.getcwd(), tempfile.gettempdir(),
               os.path.expanduser(os.path.join("~", ".workbuddy"))]
    for a in allowed:
        try:
            if os.path.commonpath([ap, os.path.abspath(a)]) == os.path.abspath(a):
                return ap
        except ValueError:
            continue
    raise ValueError(f"输出路径越界（仅允许工作目录/临时目录）：{path}")


def norm(text):
    """diff 用归一化：去空白、统一全半角标点、去格式噪声。"""
    t = re.sub(r"\s+", "", text or "")
    table = str.maketrans("（）：；，．＿×", "()：;,._x")
    return t.translate(table)


def cn_to_num(cn):
    """人民币大写 → 数值（支持 万/亿/元角分）。无法解析返回 None。"""
    if not cn:
        return None
    s = re.sub(r"^人民币", "", cn.strip())
    s = re.sub(r"[整正]$", "", s)
    m = re.search(r"([零壹贰叁肆伍陆柒捌玖拾佰仟]+)亿", s)
    total, rest = 0, s
    if m:
        total += _cn_under_yi(m.group(1)) * 10 ** 8
        rest = s[m.end():]
    parts = re.split(r"万", rest)
    if len(parts) == 2:
        total += _cn_under_yi(parts[0]) * 10 ** 4
        tail = parts[1]
    else:
        tail = parts[0]
    tail = re.sub(r"^圆|^元", "", tail)
    num = 0
    yuan_part = re.split(r"角|分", tail)[0]
    yuan_part = re.sub(r"元$", "", yuan_part)
    cur, pending = 0, 0
    for ch in yuan_part:
        if ch in CN_DIG:
            pending = CN_DIG[ch]
        elif ch in CN_UNIT:
            cur += (pending or 1) * CN_UNIT[ch]
            pending = 0
    cur += pending
    if re.search(r"[零壹贰叁肆伍陆柒捌玖]元", tail):
        num = cur
    else:
        num = cur
    total += num
    jiao = re.search(r"([零壹贰叁肆伍陆柒捌玖])角", tail)
    fen = re.search(r"([零壹贰叁肆伍陆柒捌玖])分", tail)
    if jiao:
        total += CN_DIG[jiao.group(1)] * 0.1
    if fen:
        total += CN_DIG[fen.group(1)] * 0.01
    return total if total else None


def _cn_under_yi(s):
    cur, pending = 0, 0
    for ch in s:
        if ch in CN_DIG:
            pending = CN_DIG[ch]
        elif ch in CN_UNIT:
            cur += (pending or 1) * CN_UNIT[ch]
            pending = 0
    return cur + pending


def parse_amount(txt):
    """数字金额解析：严格拒绝 修饰语/负号/非数。返回 float 或 None。"""
    if txt is None:
        return None
    s = str(txt).strip()
    if re.search(r"约|左右|大约|预计|不低于|不低于|含税|不含税|人民币|¥|￥|元", s):
        s = re.sub(r"人民币|¥|￥|元|（[^）]*）|\([^)]*\)", "", s)
        if re.search(r"约|左右|大约|预计|不低于", s):
            return None
    s = s.replace(",", "").replace("，", "")
    try:
        v = float(s)
    except ValueError:
        return None
    return v if v >= 0 else None


PLACEHOLDER_RE = re.compile(r"_{2,}|＿{2,}|×{3,}|【[^】]{0,30}】|（\s*）")
# 原始模板占位（与 fill_plan 槽位口径一致）：下划线/×××/【】/（ ）/冒号空白槽
RAW_PH_RE = re.compile(r"_{2,}|＿{2,}|×{3,}|【[^】]{0,20}】|（\s*）|\(\s*\)|[：:][ \t\u3000]{2,}(?=[（(年元天日%]|$)")


def iter_sections(content):
    """兼容 str 段落与 dict 段落；yield (section, paragraph_dict)。"""
    for sec in content.get("document", {}).get("sections", []):
        for p in sec.get("paragraphs", []):
            if isinstance(p, str):
                yield sec, {"text": p}
            else:
                yield sec, p


def sec_text(sec):
    out = []
    for p in sec.get("paragraphs", []):
        out.append(p if isinstance(p, str) else p.get("text", ""))
    return "\n".join(out)
