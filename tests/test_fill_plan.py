# -*- coding: utf-8 -*-
"""fill_plan.match_sources 置信度分级回归测试。

核心回归点（v0.3.1 修复）：
- high(auto) 仅当画像键名与槽 label 归一化后【精确一致】且值唯一；
- 仅子串命中的单一候选【不得】判 high（旧逻辑会把「名称→项目名称」串值自动填入）；
- 子串命中 / 多值冲突 → medium（ask）；无来源 → low。
"""
import json

import fill_plan as fp


def _slot(label):
    return {"id": "P0#0", "type": "paragraph", "label": label, "current": "____"}


def _profile(tmp_path, data, name="profile.json"):
    p = tmp_path / name
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return str(p)


class TestConfidence:
    def test_精确键唯一_high(self, tmp_path):
        src = _profile(tmp_path, {"项目名称": "淄博油库工程", "公司名称": "雅合科技"})
        slots = fp.match_sources([_slot("项目名称：")], [src])
        s = slots[0]
        assert s["confidence"] == "high"
        assert s["status"] == "auto"
        assert s["candidates"][0]["value"] == "淄博油库工程"

    def test_串值回归_子串单一候选降medium(self, tmp_path):
        # 旧逻辑：画像只有键「名称」，槽「项目名称」子串命中 → high → 错填公司名
        src = _profile(tmp_path, {"名称": "雅合建设有限公司"})
        slots = fp.match_sources([_slot("项目名称：")], [src])
        assert slots[0]["confidence"] == "medium"
        assert slots[0]["status"] == "ask"

    def test_精确键多值冲突_medium(self, tmp_path):
        s1 = _profile(tmp_path, {"工期": "240"}, "a.json")
        s2 = _profile(tmp_path, {"工期": "46"}, "b.json")
        slots = fp.match_sources([_slot("工期：")], [s1, s2])
        assert slots[0]["confidence"] == "medium"

    def test_无来源_low(self, tmp_path):
        src = _profile(tmp_path, {"公司名称": "雅合科技"})
        slots = fp.match_sources([_slot("项目负责人联系电话：")], [src])
        assert slots[0]["confidence"] == "low"

    def test_嵌套键末段精确匹配_high(self, tmp_path):
        src = _profile(tmp_path, {"企业": {"名称": "雅合建设有限公司"}})
        slots = fp.match_sources([_slot("名称：")], [src])
        assert slots[0]["confidence"] == "high"
        assert slots[0]["candidates"][0]["value"] == "雅合建设有限公司"

    def test_两个嵌套同名键多值_medium(self, tmp_path):
        src = _profile(tmp_path, {"企业": {"名称": "雅合A"}, "项目": {"名称": "淄博B"}})
        slots = fp.match_sources([_slot("名称：")], [src])
        assert slots[0]["confidence"] == "medium"

    def test_画像占位值过滤_low(self, tmp_path):
        src = _profile(tmp_path, {"工期": "【待补：×××】"})
        slots = fp.match_sources([_slot("工期：")], [src])
        assert slots[0]["confidence"] == "low"

    def test_label剥尾冒号(self, tmp_path):
        # 真实槽位 label 是占位符前的引导文本（不含占位符本身），如「项目经理：」
        src = _profile(tmp_path, {"项目经理": "张三"})
        slots = fp.match_sources([_slot("项目经理：")], [src])
        assert slots[0]["confidence"] == "high"

    def test_bool与空值不入候选(self, tmp_path):
        src = _profile(tmp_path, {"是否联合体": True, "备注": "", "工期": "120"})
        slots = fp.match_sources([_slot("是否联合体："), _slot("备注：")], [src])
        assert slots[0]["confidence"] == "low"
        assert slots[1]["confidence"] == "low"

    def test_编号前缀endswith精确_high(self, tmp_path):
        # 真实槽位 label 常带编号前缀（（九）工期：），视为精确命中
        src = _profile(tmp_path, {"工期": "240天"})
        slots = fp.match_sources([_slot("（九）工期：")], [src])
        assert slots[0]["confidence"] == "high"

    def test_非编号前缀endswith不算精确(self, tmp_path):
        # 「项目名称」以泛词键「名称」结尾，但前缀非编号 → 不得精确 → medium
        src = _profile(tmp_path, {"名称": "雅合建设有限公司"})
        slots = fp.match_sources([_slot("项目名称：")], [src])
        assert slots[0]["confidence"] == "medium"

    def test_整键存在时取整键(self, tmp_path):
        # 画像同时含「名称」与「项目名称」时，项目名称槽应命中整键值
        src = _profile(tmp_path, {"名称": "雅合建设有限公司", "项目名称": "淄博油库工程"})
        slots = fp.match_sources([_slot("项目名称：")], [src])
        assert slots[0]["confidence"] == "high"
        assert slots[0]["candidates"][0]["value"] == "淄博油库工程"


class TestBuildSlotsExempt:
    def test_格式标题豁免需docx环境(self):
        # build_slots 依赖 python-docx；此处仅验证豁免判定函数可用（真值测在 test_common）
        import _common
        assert _common.is_exempt_ph("【工程量清单计价】")
