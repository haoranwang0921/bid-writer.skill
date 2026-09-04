#!/usr/bin/env python3
"""smart_fill.py — 模型检索填空的任务生成与证据关卡。

本脚本不绑定任何模型或数据库 SDK。运行 bid-writer 的 agent 读取 prepare
产物，使用当前环境可用的只读数据库/知识库工具检索，再按协议提交 decisions。
validate 只放行「唯一值 + 可信来源类型 + 可定位证据」的 high 决策。

CLI:
  python smart_fill.py prepare fill_plan.json [-c data_sources.json] -o model_tasks.json
  python smart_fill.py validate fill_plan.json model_decisions.json \
      [-c data_sources.json] -o smart_fill.json [--questions smart_questions.json]
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from _common import atomic_write_json, guard_out  # noqa: E402


PROTOCOL = "bid-writer-smart-fill/v1"
ALLOWED_SOURCE_TYPES = {
    "structured",      # 企业画像/业务数据库等结构化事实
    "tender",          # 当前招标文件解析结果
    "confirmed",       # 已由用户确认并固化的事实
    "document",        # 可定位到页/节/段的原始文档
    "knowledge_base",  # 语义检索片段
}
AUTO_SOURCE_TYPES = {"structured", "tender", "confirmed"}
CRITICAL_RE = re.compile(
    r"金额|报价|单价|总价|税率|折扣|日期|时间|工期|有效期|证书|资质|许可|"
    r"信用代码|身份证|银行|账号|保证金|法定代表人"
)


def _read_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _text(value):
    return "" if value is None else str(value).strip()


def classify_slot(slot):
    """标记事实敏感度，供 agent 决定检索深度与复核方式。"""
    joined = f"{slot.get('label', '')} {slot.get('context', '')}"
    return "critical_fact" if CRITICAL_RE.search(joined) else "general_fact"


def _catalog_from_plan(plan):
    catalog = plan.get("meta", {}).get("source_catalog", [])
    return catalog if isinstance(catalog, list) else []


def prepare_bundle(plan, catalog=None, plan_path="", include_auto=False):
    """把槽位转成供模型逐项检索的任务包。"""
    sources = catalog if catalog is not None else _catalog_from_plan(plan)
    tasks = []
    for slot in plan.get("slots", []):
        if not include_auto and slot.get("status") != "ask":
            continue
        label = _text(slot.get("label"))
        context = _text(slot.get("context"))
        tasks.append({
            "slot_id": slot.get("id"),
            "slot_type": slot.get("type"),
            "label": label,
            "context": context,
            "current": slot.get("current", ""),
            "safety_class": classify_slot(slot),
            "existing_candidates": slot.get("candidates", []),
            "retrieval_goal": (
                f"确定「{label or context}」应填写的当前项目唯一事实；"
                "优先查询结构化企业库和当前招标解析结果，返回值及可复查定位。"
            ),
        })
    return {
        "meta": {
            "protocol": PROTOCOL,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "fill_plan": plan_path,
            "task_count": len(tasks),
        },
        "sources": sources,
        "tasks": tasks,
        "decision_contract": {
            "required": ["slot_id", "confidence", "query", "reason", "evidence"],
            "optional": ["canonical_field", "selected_value"],
            "evidence_required": ["value", "source", "source_type", "locator"],
            "confidence": ["high", "medium", "low"],
            "source_type": sorted(ALLOWED_SOURCE_TYPES),
        },
    }


def _source_registry(plan, catalog=None):
    sources = catalog if catalog is not None else _catalog_from_plan(plan)
    registry = {}
    for item in sources:
        if not isinstance(item, dict):
            continue
        sid = _text(item.get("id"))
        source_type = _text(item.get("source_type")).lower()
        if sid and source_type in ALLOWED_SOURCE_TYPES:
            registry[sid] = item
    return registry


def _clean_evidence(raw, slot_id, warnings, source_registry):
    if not isinstance(raw, list):
        warnings.append(f"{slot_id}: evidence 必须是数组")
        return []
    out, seen = [], set()
    for i, item in enumerate(raw):
        if not isinstance(item, dict):
            warnings.append(f"{slot_id}: evidence[{i}] 不是对象，已忽略")
            continue
        value = _text(item.get("value"))
        source = _text(item.get("source"))
        source_type = _text(item.get("source_type")).lower()
        locator = _text(item.get("locator"))
        if not value or not source or not locator:
            warnings.append(f"{slot_id}: evidence[{i}] 缺 value/source/locator，已忽略")
            continue
        if source_type not in ALLOWED_SOURCE_TYPES:
            warnings.append(f"{slot_id}: evidence[{i}] source_type={source_type!r} 不受信，已忽略")
            continue
        registered = source in source_registry
        if not registered:
            warnings.append(f"{slot_id}: evidence[{i}] source={source!r} 未登记，只能作为待确认候选")
        elif _text(source_registry[source].get("source_type")).lower() != source_type:
            registered = False
            warnings.append(f"{slot_id}: evidence[{i}] source_type 与数据源目录不一致，只能待确认")
        key = (value, source, source_type, locator)
        if key in seen:
            continue
        seen.add(key)
        out.append({
            "value": value,
            "source": source,
            "source_type": source_type,
            "locator": locator,
            "registered": registered,
            **({"record_id": _text(item.get("record_id"))} if item.get("record_id") else {}),
            **({"updated_at": _text(item.get("updated_at"))} if item.get("updated_at") else {}),
        })
    return out


def _question(slot, confidence, evidence, basis, decision=None):
    options, seen = [], set()
    for item in evidence:
        value = item["value"]
        if value in seen:
            continue
        seen.add(value)
        options.append(
            f"{value}（依据：{item['source']}；定位：{item['locator']}）"
        )
        if len(options) == 4:
            break
    if confidence == "low" or not options:
        options.append("自行填写")
    return {
        "id": slot.get("id"),
        "label": slot.get("label", ""),
        "context": slot.get("context", ""),
        "confidence": confidence,
        "options": options,
        "basis": basis,
        "canonical_field": (decision or {}).get("canonical_field", ""),
        "model_reason": (decision or {}).get("reason", ""),
        "evidence": evidence,
    }


def validate_bundle(plan, raw_decisions, catalog=None):
    """验证模型决策；任何证据不足都降级为提问，不抛给 Word 写回。"""
    slots = {s.get("id"): s for s in plan.get("slots", []) if s.get("id")}
    source_registry = _source_registry(plan, catalog)
    incoming = raw_decisions.get("decisions", raw_decisions)
    errors, warnings = [], []
    if not isinstance(incoming, list):
        incoming = []
        errors.append("decisions 必须是数组")

    by_id, duplicate_ids = {}, set()
    for raw in incoming:
        if not isinstance(raw, dict):
            errors.append("存在非对象 decision，已忽略")
            continue
        sid = _text(raw.get("slot_id"))
        if sid not in slots:
            errors.append(f"未知槽位 {sid or '<empty>'}，已忽略")
            continue
        if sid in by_id:
            duplicate_ids.add(sid)
            errors.append(f"槽位 {sid} 出现重复 decision，禁止自动填入")
            continue
        by_id[sid] = raw

    auto_fills, validated, questions, unresolved = {}, [], [], []
    for sid, slot in slots.items():
        # 原有 deterministic auto 由 fill_plan.apply 处理；模型只补 semantic ask 槽位。
        if slot.get("status") != "ask":
            continue
        raw = by_id.get(sid)
        if raw is None or sid in duplicate_ids:
            existing = []
            for cand in slot.get("candidates", []):
                value = _text(cand.get("value")) if isinstance(cand, dict) else ""
                if value:
                    existing.append({
                        "value": value,
                        "source": _text(cand.get("source")) or "fill_plan",
                        "source_type": "structured",
                        "locator": "原始字段匹配候选（待语义确认）",
                    })
            conf = "medium" if existing else "low"
            basis = "模型未返回唯一可验证决策" if raw is None else "模型返回了重复决策"
            questions.append(_question(slot, conf, existing, basis))
            unresolved.append(sid)
            continue

        model_conf = _text(raw.get("confidence")).lower()
        if model_conf not in {"high", "medium", "low"}:
            warnings.append(f"{sid}: confidence={model_conf!r} 非法，按 low 处理")
            model_conf = "low"
        selected = _text(raw.get("selected_value"))
        evidence = _clean_evidence(raw.get("evidence", []), sid, warnings, source_registry)
        distinct_values = {e["value"] for e in evidence}
        selected_evidence = [e for e in evidence if e["value"] == selected]
        trusted_selected = [e for e in selected_evidence
                            if e["source_type"] in AUTO_SOURCE_TYPES and e.get("registered")]
        issues = []
        auto_eligible = False
        effective = model_conf

        if model_conf == "high":
            if not selected:
                issues.append("high 决策缺 selected_value")
            if selected and not selected_evidence:
                issues.append("selected_value 未出现在证据值中")
            if len(distinct_values) != 1:
                issues.append("证据存在零个或多个不同值")
            if selected_evidence and not trusted_selected:
                issues.append("选中值仅来自文档/语义知识库，不允许自动填入")
            if not _text(raw.get("query")):
                issues.append("缺检索 query")
            if not _text(raw.get("reason")):
                issues.append("缺决策 reason")
            if issues:
                effective = "medium" if evidence else "low"
            else:
                auto_eligible = True
                effective = "high"
                auto_fills[sid] = selected
        elif model_conf == "medium":
            effective = "medium" if evidence else "low"
        else:
            effective = "low"

        normalized = {
            "slot_id": sid,
            "canonical_field": _text(raw.get("canonical_field")),
            "query": _text(raw.get("query")),
            "selected_value": selected,
            "model_confidence": model_conf,
            "effective_confidence": effective,
            "reason": _text(raw.get("reason")),
            "evidence": evidence,
            "validation": {
                "auto_eligible": auto_eligible,
                "safety_class": classify_slot(slot),
                "issues": issues,
            },
        }
        validated.append(normalized)
        if not auto_eligible:
            basis = "；".join(issues) or "模型未判定为唯一高置信度事实"
            questions.append(_question(slot, effective, evidence, basis, normalized))
            unresolved.append(sid)

    return {
        "meta": {
            "protocol": PROTOCOL,
            "validated": True,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "stats": {
                "decisions": len(validated),
                "auto": len(auto_fills),
                "unresolved": len(unresolved),
                "errors": len(errors),
                "warnings": len(warnings),
            },
            "errors": errors,
            "warnings": warnings,
        },
        "auto_fills": auto_fills,
        "decisions": validated,
        "questions": questions,
        "unresolved": unresolved,
    }


def _load_catalog(path):
    if not path:
        return None
    data = _read_json(path)
    sources = data.get("sources", data) if isinstance(data, dict) else data
    if not isinstance(sources, list):
        raise ValueError("数据源目录必须是数组或包含 sources 数组")
    return sources


def cmd_prepare(args):
    plan = _read_json(args.plan)
    bundle = prepare_bundle(
        plan,
        catalog=_load_catalog(args.catalog),
        plan_path=args.plan,
        include_auto=args.include_auto,
    )
    atomic_write_json(guard_out(args.out), bundle)
    print(f"OK model_tasks={len(bundle['tasks'])} -> {args.out}")


def cmd_validate(args):
    plan = _read_json(args.plan)
    raw = _read_json(args.decisions)
    result = validate_bundle(plan, raw, catalog=_load_catalog(args.catalog))
    atomic_write_json(guard_out(args.out), result)
    if args.questions:
        atomic_write_json(guard_out(args.questions), {"questions": result["questions"]})
    stats = result["meta"]["stats"]
    print(
        f"OK model_auto={stats['auto']} unresolved={stats['unresolved']} "
        f"errors={stats['errors']} warnings={stats['warnings']} -> {args.out}"
    )


def main():
    ap = argparse.ArgumentParser(description="模型检索填空的任务生成与证据关卡")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("prepare")
    p1.add_argument("plan", help="fill_plan.json")
    p1.add_argument("-c", "--catalog", default="", help="可选 data_sources.json")
    p1.add_argument("-o", "--out", default="model_tasks.json")
    p1.add_argument("--include-auto", action="store_true",
                    help="同时输出原 deterministic auto 槽位供模型审计")
    p1.set_defaults(fn=cmd_prepare)
    p2 = sub.add_parser("validate")
    p2.add_argument("plan", help="fill_plan.json")
    p2.add_argument("decisions", help="模型生成的 model_decisions.json")
    p2.add_argument("-c", "--catalog", default="", help="prepare 时使用的同一数据源目录")
    p2.add_argument("-o", "--out", default="smart_fill.json")
    p2.add_argument("--questions", default="", help="另存待确认问题单")
    p2.set_defaults(fn=cmd_validate)
    args = ap.parse_args()
    try:
        args.fn(args)
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
