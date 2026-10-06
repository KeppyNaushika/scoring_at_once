"""記録した正解 (tests/golden/) と, 今のコードの挙動を比べる.

正解は作り直しの前 (commit 99feca3) の挙動を記録したもの. その後, 意図して直した不具合の分だけ取り直した:
- 採点済み答案の合計点に, 大問が未設定の設問の点数も含める (export:*, preset1, preset2 の答案 0・4・5)

正解を取り直すとき (挙動を意図して変えたとき) は:
    for s in roster grading area projects sao preset1 preset2 export:default export:corners export:edges; do
        python tests/capture.py . $s tests/golden/${s/:/_}.json
    done
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).parent.parent
SCENARIOS = [
    "roster", "grading", "area", "projects", "sao", "preset1", "preset2",
    "export:default", "export:corners", "export:edges",
]


@pytest.mark.parametrize("scenario", SCENARIOS)
def test_same_as_golden(scenario: str, tmp_path: Path) -> None:
    output = tmp_path / "result.json"
    subprocess.run(
        [sys.executable, str(ROOT / "tests" / "capture.py"), str(ROOT), scenario, str(output)],
        check=True, timeout=120,
    )
    actual = json.loads(output.read_text(encoding="utf-8"))
    expected = json.loads((ROOT / "tests" / "golden" / f"{scenario.replace(':', '_')}.json").read_text(encoding="utf-8"))
    assert actual.get("exception") is None
    assert actual["errors"] == []
    # ダイアログの文言は比べない. 採点済み答案画像は点数の文字のフォントが OS ごとに違う (メイリオ/ヒラギノ) ので,
    # 正解を記録した macOS でだけ画素を比べる (他の OS でも書き出しの処理自体は通す)
    ignored = {"messages"} | (set() if sys.platform == "darwin" else {"png"})
    for key in expected.keys() - ignored:
        assert actual.get(key) == expected[key], f"{scenario}: {key} が変わっています"
