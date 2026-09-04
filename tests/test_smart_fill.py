# -*- coding: utf-8 -*-
"""模型检索填空的证据关卡回归测试。"""
import json
import os
import subprocess
import sys
from pathlib import Path

import fill_plan
import smart_fill
from docx import Document


def _slot(sid="P0#0", label="法定代表人：", candidates=None, status="ask"):
    return {
        "id": sid,
        "type": "paragraph",
        "label": label,
        "context": f"{label}____",
        "current": "____",
        "candidates": candidates or [],
        "confidence": "low" if not candidates else "medium",
        "status": status,
    }


def _plan(*slots):
    return {"meta": {"source_catalog": [{
        "id": "company", "kind": "json", "location": "profile.json",
        "source_type": "structured", "access": "read_only",
    }]}, "slots": list(slots)}


def _decision(value="张三", source_type="structured", confidence="high"):
    return {
        "slot_id": "P0#0",
        "canonical_field": "basic.legal_representative",
        "query": "查询当前投标主体法定代表人",
        "selected_value": value,
        "confidence": confidence,
        "reason": "当前投标主体只有一个有效记录",
        "evidence": [{
            "value": value,
            "source": "company",
            "source_type": source_type,
            "locator": "companies/1/basic/legal_representative",
        }],
    }


class TestPrepare:
    def test_只生成ask任务并携带数据源(self):
        plan = _plan(
            _slot(),
            _slot("P1#0", "项目名称：", [{"value": "项目A", "source": "tender.json"}], "auto"),
        )
        bundle = smart_fill.prepare_bundle(plan, plan_path="fill_plan.json")
        assert bundle["meta"]["protocol"] == smart_fill.PROTOCOL
        assert [t["slot_id"] for t in bundle["tasks"]] == ["P0#0"]
        assert bundle["tasks"][0]["safety_class"] == "critical_fact"
        assert bundle["sources"][0]["id"] == "company"


class TestValidate:
    def test_唯一结构化证据允许自动填(self):
        result = smart_fill.validate_bundle(_plan(_slot()), {"decisions": [_decision()]})
        assert result["auto_fills"] == {"P0#0": "张三"}
        assert result["questions"] == []
        assert result["decisions"][0]["validation"]["auto_eligible"] is True

    def test_知识库证据即使模型报high也降medium(self):
        result = smart_fill.validate_bundle(
            _plan(_slot()), {"decisions": [_decision(source_type="knowledge_base")]}
        )
        assert result["auto_fills"] == {}
        assert result["questions"][0]["confidence"] == "medium"
        assert "不允许自动填入" in result["questions"][0]["basis"]

    def test_未登记来源冒充结构化数据也不能自动填(self):
        d = _decision()
        d["evidence"][0]["source"] = "unknown-db"
        result = smart_fill.validate_bundle(_plan(_slot()), {"decisions": [d]})
        assert result["auto_fills"] == {}
        assert result["questions"][0]["confidence"] == "medium"
        assert result["decisions"][0]["evidence"][0]["registered"] is False

    def test_冲突值禁止自动填(self):
        d = _decision()
        d["evidence"].append({
            "value": "李四",
            "source": "project-db",
            "source_type": "structured",
            "locator": "companies/2/basic/legal_representative",
        })
        result = smart_fill.validate_bundle(_plan(_slot()), {"decisions": [d]})
        assert result["auto_fills"] == {}
        assert result["questions"][0]["confidence"] == "medium"
        assert len(result["questions"][0]["options"]) == 2

    def test_选中值不在证据中禁止自动填(self):
        d = _decision(value="张三")
        d["selected_value"] = "模型编造值"
        result = smart_fill.validate_bundle(_plan(_slot()), {"decisions": [d]})
        assert result["auto_fills"] == {}
        assert "未出现在证据值" in result["questions"][0]["basis"]

    def test_模型漏答回退到原候选提问(self):
        slot = _slot(candidates=[{"value": "项目A", "source": "tender.json"}])
        result = smart_fill.validate_bundle(_plan(slot), {"decisions": []})
        assert result["auto_fills"] == {}
        assert result["questions"][0]["confidence"] == "medium"
        assert "项目A" in result["questions"][0]["options"][0]


class TestApplyContract:
    def test_只接受验证产物并保留审计(self, tmp_path):
        plan = _plan(_slot())
        result = smart_fill.validate_bundle(plan, {"decisions": [_decision()]})
        path = tmp_path / "smart.json"
        path.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
        auto, audit = fill_plan.load_validated_smart(str(path))
        assert auto == {"P0#0": "张三"}
        assert audit["P0#0"]["origin"] == "model_evidence"
        assert audit["P0#0"]["evidence"][0]["locator"].startswith("companies/")

    def test_拒绝伪造未验证文件(self, tmp_path):
        path = tmp_path / "fake.json"
        path.write_text(json.dumps({"auto_fills": {"P0#0": "假值"}}), encoding="utf-8")
        try:
            fill_plan.load_validated_smart(str(path))
        except ValueError as exc:
            assert "受控决策文件" in str(exc)
        else:
            raise AssertionError("未验证 smart 文件不应被接受")

    def test_cli_prepare_validate往返(self, tmp_path):
        plan_path = tmp_path / "plan.json"
        decisions_path = tmp_path / "decisions.json"
        tasks_path = tmp_path / "tasks.json"
        resolved_path = tmp_path / "resolved.json"
        plan_path.write_text(json.dumps(_plan(_slot()), ensure_ascii=False), encoding="utf-8")
        decisions_path.write_text(
            json.dumps({"decisions": [_decision()]}, ensure_ascii=False), encoding="utf-8"
        )
        script = Path(__file__).parents[1] / "scripts" / "smart_fill.py"
        child_env = os.environ.copy()
        child_env["PYTHONUTF8"] = "1"
        prepared = subprocess.run(
            [sys.executable, str(script), "prepare", str(plan_path), "-o", str(tasks_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=child_env,
        )
        assert prepared.returncode == 0, prepared.stderr
        validated = subprocess.run(
            [sys.executable, str(script), "validate", str(plan_path), str(decisions_path),
             "-o", str(resolved_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=child_env,
        )
        assert validated.returncode == 0, validated.stderr
        assert json.loads(resolved_path.read_text(encoding="utf-8"))["auto_fills"] == {
            "P0#0": "张三"
        }

    def test_apply将模型值写入fills并允许人工覆盖(self, tmp_path):
        template_path = tmp_path / "template.docx"
        doc = Document()
        doc.add_paragraph("法定代表人：____")
        doc.save(str(template_path))
        plan_path = tmp_path / "plan.json"
        smart_path = tmp_path / "smart.json"
        fills_path = tmp_path / "fills.json"
        plan = _plan(_slot())
        plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
        smart = smart_fill.validate_bundle(plan, {"decisions": [_decision()]})
        smart_path.write_text(json.dumps(smart, ensure_ascii=False), encoding="utf-8")
        answers_path = tmp_path / "answers.json"
        answers_path.write_text(
            json.dumps({"answers": {"P0#0": "李四"}}, ensure_ascii=False), encoding="utf-8"
        )
        fill_plan.cmd_apply(type("Args", (), {
            "template": str(template_path),
            "plan": str(plan_path),
            "smart": str(smart_path),
            "answers": str(answers_path),
            "out": str(fills_path),
        })())
        fills = json.loads(fills_path.read_text(encoding="utf-8"))
        assert fills["fills"]["P0#0"] == "李四"
        assert fills["audit"]["P0#0"]["origin"] == "human_answer"
