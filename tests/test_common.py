# -*- coding: utf-8 -*-
"""_common 共享工具单元测试。"""
import os
import tempfile

import pytest

from _common import (RAW_PH_RE, cn_to_num, guard_out, is_exempt_ph, norm,
                     parse_amount)


class TestCnToNum:
    def test_整数元(self):
        assert cn_to_num("壹仟贰佰元整") == 1200

    def test_万元(self):
        assert cn_to_num("壹万贰仟元整") == 12000

    def test_拾元零头(self):
        assert cn_to_num("拾元整") == 10

    def test_角分(self):
        assert cn_to_num("壹佰贰拾叁元肆角伍分") == pytest.approx(123.45)

    def test_整元无角分(self):
        assert cn_to_num("玖佰捌拾柒元整") == 987

    def test_亿级(self):
        assert cn_to_num("贰亿零伍佰万元整") == 205000000

    def test_人民币前缀(self):
        assert cn_to_num("人民币伍佰元整") == 500

    def test_无法解析(self):
        assert cn_to_num(None) is None
        assert cn_to_num("") is None


class TestParseAmount:
    def test_纯数字(self):
        assert parse_amount("12345") == 12345.0

    def test_带元字(self):
        assert parse_amount("12345元") == 12345.0

    def test_千分位(self):
        assert parse_amount("12,345.67") == 12345.67

    def test_模糊修饰语拒收(self):
        assert parse_amount("约1000") is None
        assert parse_amount("1000左右") is None
        assert parse_amount("不低于1000") is None
        assert parse_amount("大约1000元") is None

    def test_负数拒收(self):
        assert parse_amount("-5") is None

    def test_非数(self):
        assert parse_amount("abc") is None
        assert parse_amount(None) is None


class TestNorm:
    def test_去空白统一括号(self):
        assert norm("项目 名称：（甲）") == "项目名称:(甲)"


class TestExemptPh:
    def test_格式标题豁免(self):
        assert is_exempt_ph("【工程量清单计价】") is True

    def test_可填占位不豁免(self):
        assert is_exempt_ph("【待填：×××】") is False
        assert is_exempt_ph("____") is False
        assert is_exempt_ph("【其他格式词】") is False

    def test_raw_ph_命中格式标题(self):
        m = next(RAW_PH_RE.finditer("投标函【工程量清单计价】"))
        assert is_exempt_ph(m.group(0)) is True


class TestGuardOut:
    def test_工作目录允许(self, monkeypatch):
        allowed = tempfile.gettempdir()
        monkeypatch.chdir(os.path.dirname(allowed))
        p = os.path.join(allowed, "out.md")
        assert guard_out(p) == os.path.abspath(p)

    def test_越界拒绝(self):
        bad = "C:/Users/Public/evil.md" if os.name == "nt" else "/etc/evil.md"
        with pytest.raises(ValueError):
            guard_out(bad)
