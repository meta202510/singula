"""Offline-only test runner. This module never constructs a network client."""

from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import tempfile

from novel_critique import CritiqueError, evaluate, parse_response

MARKER = "【モックテスト専用・本番審査ではありません】"


class MockStopped(CritiqueError):
    pass


@dataclass(frozen=True)
class MockOptions:
    resume: bool = False
    stop_after: int | None = None
    fail_at: int | None = None

    def validate(self):
        for value in (self.stop_after, self.fail_at):
            if value is not None and (type(value) is not int or not 1 <= value <= 6):
                raise CritiqueError("モックの停止・エラー位置は1～6で指定してください。")


def state_path(output: Path) -> Path:
    return output.with_name(output.name + ".mock-state.json")


def write_state(path: Path, data: dict):
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, prefix=".mock-", delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(data, stream, ensure_ascii=False)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def evaluate_mock(novel: str, filename: str, output: Path, options: MockOptions, progress):
    options.validate()
    path = state_path(output)
    fingerprint = hashlib.sha256((filename + "\0" + novel).encode("utf-8")).hexdigest()
    state = {"mode": "singula-mock-v1", "fingerprint": fingerprint, "results": []}
    if options.resume:
        try:
            state = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            raise CritiqueError("モック再開ファイルを読み取れません。保存先とファイルを確認してください。") from None
        if (not isinstance(state, dict) or state.get("mode") != "singula-mock-v1"
                or state.get("fingerprint") != fingerprint
                or not isinstance(state.get("results"), list) or len(state["results"]) > 6
                or any(not isinstance(item, str) or not item.startswith(MARKER) for item in state["results"])):
            raise CritiqueError("モック再開データが不正、または入力ファイルが変更されています。")
    elif os.path.lexists(path):
        raise CritiqueError("モックの途中データが存在します。再開を選ぶか、別の保存先を指定してください。")
    else:
        write_state(path, state)

    index = 0

    def review(instructions, text):
        nonlocal index
        index += 1
        if index <= len(state["results"]):
            return state["results"][index - 1]
        if options.fail_at == index:
            raise CritiqueError(f"[モック] {index}回目の評価でテスト用エラーを発生させました。通信・課金はありません。")
        if index <= 5:
            result = f"{MARKER}\n審査員{index}のダミー批評です。構成には読みやすさがあり、人物描写には掘り下げの余地があります。これは動作確認用の固定文で、本文への実際の評価ではありません。"
        else:
            result = f"{MARKER}\n5人のダミー批評に基づく総評です。長所と改善点を整理する処理のテスト用であり、実際の審査結果ではありません。"
        result = parse_response({"status": "completed", "output": [{
            "type": "message", "role": "assistant", "content": [{"type": "output_text", "text": result}],
        }]})
        state["results"].append(result)
        write_state(path, state)
        if options.stop_after == index:
            raise MockStopped(f"[モック] {index}/6 完了後に停止しました。途中データ: {path}")
        return result

    return evaluate(novel, review, lambda text: progress("[モック] " + text))
