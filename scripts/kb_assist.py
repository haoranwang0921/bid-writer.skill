#!/usr/bin/env python3
"""kb_assist.py — 环节二知识库检索辅助（可选步骤）

对 fill_plan.json 中所有 status=ask 的槽位（medium/low 档），
调用本地检索服务获取参考知识片段，写入每个槽位的 "kb_hint" 字段。
agent 在提问关口可把 kb_hint 作为背景参考，帮助用户判断填写依据。

用法：
    python kb_assist.py <fill_plan.json> -o fill_plan_kb.json
    python kb_assist.py <fill_plan.json> --service http://127.0.0.1:8765 --top-k 3

检索服务不可达时静默跳过（不阻断流程），输出与输入内容相同。

约束：
- kb_hint 是"参考背景"，不是候选值，不得作为 auto 填入依据（W2 不变）。
- medium 槽位：kb_hint 帮助解释候选值的来源背景；
- low 槽位：kb_hint 提供填写方向参考，最终仍须用户确认。
"""
import argparse
import json
import os
import sys
import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(__file__))
from _common import atomic_write_json, guard_out  # noqa: E402


def query_service(service_url: str, query: str, top_k: int = 2) -> str:
    """调用 /query 接口，返回 context_block 文本；失败返回空串。"""
    payload = {"query": query, "top_k": top_k, "format": "context"}
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        f"{service_url}/query",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=8) as resp:
            result = json.loads(resp.read().decode("utf-8"))
        return result.get("context_block", "")
    except (urllib.error.URLError, OSError, json.JSONDecodeError, ValueError):
        return ""


def build_query(slot: dict) -> str:
    """构造检索查询句：优先用 label，context 补充语义。"""
    label = (slot.get("label") or "").strip()
    context = (slot.get("context") or "").strip()
    if label and context:
        q = f"{label} {context[:40]}"
    else:
        q = label or context[:60]
    return q.strip()


def main():
    ap = argparse.ArgumentParser(description="环节二知识库检索辅助")
    ap.add_argument("fill_plan", help="fill_plan.json（scan 产物）")
    ap.add_argument("-o", "--out", default="fill_plan_kb.json",
                    help="输出路径（默认 fill_plan_kb.json）")
    ap.add_argument("--service", default="http://127.0.0.1:8765",
                    help="检索服务地址（默认 127.0.0.1:8765）")
    ap.add_argument("--top-k", type=int, default=2,
                    help="每个槽位返回的知识片段数（默认 2）")
    args = ap.parse_args()

    try:
        plan = json.load(open(args.fill_plan, encoding="utf-8"))
    except Exception as e:
        print(f"ERROR: fill_plan 读取失败: {e}", file=sys.stderr)
        sys.exit(2)

    slots = plan.get("slots", [])

    # 探测服务连通性
    svc_ok = False
    try:
        with urllib.request.urlopen(f"{args.service}/health", timeout=4) as r:
            svc_ok = (r.status == 200)
    except Exception:
        pass

    if not svc_ok:
        print(f"WARN 检索服务 {args.service} 不可达，跳过知识库辅助（不阻断流程）",
              file=sys.stderr)
        plan.setdefault("meta", {})["kb_assist"] = {"status": "skipped",
                                                      "reason": "服务不可达"}
        atomic_write_json(guard_out(args.out), plan)
        print(f"OK（无增强）-> {args.out}")
        return

    enhanced, failed = 0, 0
    for slot in slots:
        if slot.get("status") != "ask":
            continue  # high/auto 槽位无需检索辅助
        q = build_query(slot)
        if len(q) < 4:
            continue
        block = query_service(args.service, q, top_k=args.top_k)
        if block and "未检索到" not in block:
            # 截断至 300 字，避免提问时背景过长
            slot["kb_hint"] = block[:300]
            enhanced += 1
        else:
            failed += 1

    plan.setdefault("meta", {})["kb_assist"] = {
        "status": "done",
        "enhanced": enhanced,
        "no_hit": failed,
        "service": args.service,
    }
    atomic_write_json(guard_out(args.out), plan)
    print(f"OK slots={len(slots)} 检索增强={enhanced} 无命中={failed} -> {args.out}")


if __name__ == "__main__":
    main()
