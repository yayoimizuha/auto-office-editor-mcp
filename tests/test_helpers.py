"""
Office 不要な純粋ロジックのユニットテスト。

対象:
  - helpers.py : 色変換・単位変換・DataFrame変換ロジック
  - tools/preview.py : _resolve_pages のページ番号正規化ロジック
  - executor.py : execute_script の exec 実行・エラーハンドリング・stdout キャプチャ
"""

from __future__ import annotations

import math

import pandas as pd
import pytest


# ===========================================================================
# helpers.py : 色変換
# ===========================================================================


class TestRgb:
    def test_red(self):
        from auto_office_editor_mcp.helpers import rgb

        assert rgb(255, 0, 0) == 255

    def test_green(self):
        from auto_office_editor_mcp.helpers import rgb

        assert rgb(0, 255, 0) == 65280

    def test_blue(self):
        from auto_office_editor_mcp.helpers import rgb

        assert rgb(0, 0, 255) == 16711680

    def test_white(self):
        from auto_office_editor_mcp.helpers import rgb

        assert rgb(255, 255, 255) == 16777215

    def test_black(self):
        from auto_office_editor_mcp.helpers import rgb

        assert rgb(0, 0, 0) == 0


class TestHexToRgbOle:
    def test_red_with_hash(self):
        from auto_office_editor_mcp.helpers import hex_to_rgb_ole, rgb

        assert hex_to_rgb_ole("#FF0000") == rgb(255, 0, 0)

    def test_red_without_hash(self):
        from auto_office_editor_mcp.helpers import hex_to_rgb_ole, rgb

        assert hex_to_rgb_ole("FF0000") == rgb(255, 0, 0)

    def test_blue(self):
        from auto_office_editor_mcp.helpers import hex_to_rgb_ole, rgb

        assert hex_to_rgb_ole("#0000FF") == rgb(0, 0, 255)

    def test_lowercase(self):
        from auto_office_editor_mcp.helpers import hex_to_rgb_ole, rgb

        assert hex_to_rgb_ole("#ff0000") == rgb(255, 0, 0)


# ===========================================================================
# helpers.py : 単位変換
# ===========================================================================


class TestUnitConversion:
    def test_cm_to_pt(self):
        from auto_office_editor_mcp.helpers import cm_to_pt

        assert abs(cm_to_pt(1.0) - 28.3465) < 0.001

    def test_cm_to_pt_zero(self):
        from auto_office_editor_mcp.helpers import cm_to_pt

        assert cm_to_pt(0.0) == 0.0

    def test_inch_to_pt(self):
        from auto_office_editor_mcp.helpers import inch_to_pt

        assert inch_to_pt(1.0) == 72.0

    def test_pt_to_cm(self):
        from auto_office_editor_mcp.helpers import pt_to_cm

        assert abs(pt_to_cm(28.3465) - 1.0) < 0.001

    def test_cm_pt_roundtrip(self):
        from auto_office_editor_mcp.helpers import cm_to_pt, pt_to_cm

        assert abs(pt_to_cm(cm_to_pt(5.0)) - 5.0) < 0.001

    def test_inch_pt_roundtrip(self):
        from auto_office_editor_mcp.helpers import inch_to_pt, pt_to_cm

        # 1インチ = 72pt = 2.54cm
        assert abs(pt_to_cm(inch_to_pt(1.0)) - 2.54) < 0.01


# ===========================================================================
# tools/preview.py : _resolve_pages
# ===========================================================================


class TestResolvePages:
    def _call(self, total, pages, label="ページ"):
        from auto_office_editor_mcp.tools.preview import _resolve_pages

        return _resolve_pages(total, pages, label)

    def test_none_within_limit(self):
        """未指定で上限以下 → 全ページの 0始まりリスト"""
        result = self._call(5, None)
        assert result == [0, 1, 2, 3, 4]

    def test_none_at_limit(self):
        """未指定でちょうど上限（20）→ OK"""
        from auto_office_editor_mcp.tools.preview import MAX_PAGES

        result = self._call(MAX_PAGES, None)
        assert result == list(range(MAX_PAGES))

    def test_none_over_limit(self):
        """未指定で上限超え → ValueError"""
        from auto_office_editor_mcp.tools.preview import MAX_PAGES

        with pytest.raises(ValueError, match="ページ数"):
            self._call(MAX_PAGES + 1, None)

    def test_single_int_first(self):
        """int 指定（1ページ目）→ [0]"""
        assert self._call(5, 1) == [0]

    def test_single_int_last(self):
        """int 指定（最終ページ）→ [total-1]"""
        assert self._call(5, 5) == [4]

    def test_list_of_ints(self):
        """list 指定 → 対応する 0始まりインデックスリスト"""
        assert self._call(5, [1, 3, 5]) == [0, 2, 4]

    def test_single_page_list(self):
        """list で1要素 → 1要素のリスト"""
        assert self._call(5, [2]) == [1]

    def test_out_of_range_int(self):
        """範囲外の int → ValueError"""
        with pytest.raises(ValueError, match="範囲外"):
            self._call(3, 4)

    def test_out_of_range_zero(self):
        """0 は 1始まりなので範囲外 → ValueError"""
        with pytest.raises(ValueError, match="範囲外"):
            self._call(3, 0)

    def test_out_of_range_in_list(self):
        """list 内に範囲外あり → ValueError"""
        with pytest.raises(ValueError, match="範囲外"):
            self._call(3, [1, 5])

    def test_label_in_error_message(self):
        """エラーメッセージにラベルが含まれる"""
        with pytest.raises(ValueError, match="スライド"):
            self._call(25, None, label="スライド")


# ===========================================================================
# executor.py : execute_script（Office 不要部分）
# ===========================================================================


class TestExecuteScript:
    """
    AppManager.start() を呼ばない状態でも exec 自体は動く。
    xl_app / wd_app / ppt_app は None になる。
    """

    def test_simple_arithmetic(self):
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("x = 1 + 1\nprint(x)")
        assert "2" in result

    def test_success_message(self):
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("x = 42")
        assert "実行成功" in result

    def test_print_captured(self):
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("print('hello world')")
        assert "hello world" in result

    def test_print_output_section(self):
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("print('test')")
        assert "--- 出力 ---" in result

    def test_syntax_error(self):
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("def foo(:")
        assert "エラー" in result

    def test_runtime_error_zero_division(self):
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("1 / 0")
        assert "ZeroDivisionError" in result

    def test_runtime_error_name_error(self):
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("undefined_variable")
        assert "NameError" in result

    def test_pd_injected(self):
        """pandas が注入済みであること"""
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("print(pd.DataFrame({'a': [1]}).shape)")
        assert "(1, 1)" in result

    def test_np_injected(self):
        """numpy が注入済みであること"""
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("print(np.array([1, 2, 3]).sum())")
        assert "6" in result

    def test_rgb_helper_injected(self):
        """rgb() ヘルパーが注入済みであること"""
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("print(rgb(255, 0, 0))")
        assert "255" in result

    def test_cm_to_pt_helper_injected(self):
        """cm_to_pt() ヘルパーが注入済みであること"""
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("print(cm_to_pt(1.0))")
        assert "28" in result

    def test_xl_app_is_none_without_start(self):
        """モジュールレベルの manager が start() されていない場合 xl_app が None
        （integration テストと一緒に実行すると app_manager が起動済みのため
        このテストは単体実行時のみ有効）"""
        import auto_office_editor_mcp.app_manager as _am

        if _am.manager._started:
            pytest.skip("app_manager is already started by integration fixture")
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("print(xl_app is None)")
        assert "True" in result

    def test_multiple_prints(self):
        """複数の print が全て出力に含まれる"""
        from auto_office_editor_mcp.executor import execute_script

        result = execute_script("print('line1')\nprint('line2')\nprint('line3')")
        assert "line1" in result
        assert "line2" in result
        assert "line3" in result
