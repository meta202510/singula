"""Local desktop interface. Run: python novel_critique_gui.py"""

from __future__ import annotations

from dataclasses import dataclass, field
import os
from pathlib import Path
from queue import Empty, Queue
import threading

from novel_critique import (
    DEFAULT_MODEL, CritiqueError, OpenAIReviewer, evaluate, format_result, save_result,
    threshold_result, validate_thresholds,
)


@dataclass(frozen=True)
class Settings:
    input_path: Path
    output_path: Path
    api_key: str = field(repr=False)
    model: str = DEFAULT_MODEL
    encoding: str = "utf-8-sig"
    max_output_tokens: int = 4000
    timeout: float = 180
    min_chars: int | None = None
    max_chars: int | None = None


def run_job(settings: Settings, events: Queue) -> None:
    """Worker: no Tk calls; all user-visible changes travel through a queue."""
    try:
        if settings.input_path.resolve() == settings.output_path.resolve():
            raise CritiqueError("入力ファイルと異なる保存先を選択してください。")
        if os.path.lexists(settings.output_path):
            raise CritiqueError("保存先が既に存在します。別の名前を選択してください。")
        if not settings.output_path.parent.is_dir():
            raise CritiqueError("保存先のフォルダーが存在しません。")
        if not settings.model.strip():
            raise CritiqueError("モデル名を入力してください。")
        if settings.encoding not in ("utf-8-sig", "cp932", "shift_jis"):
            raise CritiqueError("入力文字コードを選択してください。")
        if not 256 <= settings.max_output_tokens <= 32000 or not 0 < settings.timeout <= 3600:
            raise CritiqueError("出力トークンは256～32000、タイムアウトは1～3600秒にしてください。")
        novel = settings.input_path.read_text(encoding=settings.encoding)
        skipped = threshold_result(settings.input_path.name, novel, settings.min_chars, settings.max_chars)
        if skipped is not None:
            save_result(settings.output_path, skipped)
            events.put(("skipped", skipped))
            return
        if not novel.strip():
            raise CritiqueError("入力ファイルが空です。")
        if not settings.api_key.strip():
            raise CritiqueError("APIキーを入力するか、環境変数 OPENAI_API_KEY を設定してください。")
        reviewer = OpenAIReviewer(
            settings.api_key.strip(), settings.model.strip(), settings.max_output_tokens, settings.timeout,
        )
        reviews, overall = evaluate(novel, reviewer.review, lambda text: events.put(("progress", text)))
        result = format_result(settings.input_path.name, reviews, overall)
        save_result(settings.output_path, result)
        events.put(("done", result))
    except UnicodeError:
        events.put(("error", "入力の文字コードを読み取れません。文字コードの設定を確認してください。"))
    except (CritiqueError, OSError) as exc:
        events.put(("error", str(exc)))
    except Exception:
        # Do not serialize arbitrary exceptions: they might contain request data.
        events.put(("error", "予期しないエラーが発生しました。設定と入力ファイルを確認してください。"))


def main() -> None:
    # Keep the worker importable and testable on machines without Tk/display support.
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText

    class Application:
        def __init__(self, root):
            self.root = root
            self.running = False
            self.events = Queue()
            self.controls = []
            self.completed = 0
            root.title("小説批評 — 5人の審査員と総評")
            root.geometry("860x820")
            root.minsize(680, 720)
            root.protocol("WM_DELETE_WINDOW", self.close)
            root.columnconfigure(0, weight=1)
            root.rowconfigure(0, weight=1)
            frame = ttk.Frame(root, padding=18)
            frame.grid(sticky="nsew")
            frame.columnconfigure(1, weight=1)
            frame.rowconfigure(16, weight=1)
            ttk.Label(frame, text="小説批評", font=("", 20, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
            ttk.Label(frame, text="5人が本文を独立して批評し、その批評だけから総評を作成します。").grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 16))
            self.input_path = tk.StringVar()
            self.output_path = tk.StringVar()
            self.api_key = tk.StringVar(value=os.environ.get("OPENAI_API_KEY", ""))
            self.model = tk.StringVar(value=DEFAULT_MODEL)
            self.encoding = tk.StringVar(value="utf-8-sig")
            self.tokens = tk.StringVar(value="4000")
            self.timeout = tk.StringVar(value="180")
            self.min_chars = tk.StringVar()
            self.max_chars = tk.StringVar()
            self.status = tk.StringVar(value="入力ファイルと保存先を選択してください。")
            self.entry(frame, 2, "小説ファイル", self.input_path, self.select_input)
            self.entry(frame, 3, "結果の保存先", self.output_path, self.select_output)
            self.entry(frame, 4, "APIキー", self.api_key, password=True)
            ttk.Label(frame, text="キーはこの画面でのみ保持し、設定ファイルには保存しません。").grid(row=5, column=1, columnspan=2, sticky="w")
            self.entry(frame, 6, "モデル", self.model)
            ttk.Label(frame, text="入力文字コード").grid(row=7, column=0, sticky="w", padx=(0, 12), pady=5)
            encoding = ttk.Combobox(frame, textvariable=self.encoding, values=("utf-8-sig", "cp932", "shift_jis"), state="readonly")
            encoding.grid(row=7, column=1, columnspan=2, sticky="ew", pady=5)
            self.controls.append((encoding, "readonly"))
            self.entry(frame, 8, "最大出力トークン／評価", self.tokens)
            self.entry(frame, 9, "通信タイムアウト（秒）", self.timeout)
            self.entry(frame, 10, "入力文字数の下限", self.min_chars)
            self.entry(frame, 11, "入力文字数の上限", self.max_chars)
            ttk.Label(frame, text="閾値は空欄で制限なし。範囲外はAPIを呼ばず「評価無し」を保存します。").grid(row=12, column=0, columnspan=3, sticky="w")
            ttk.Label(frame, text="範囲内の場合、小説全文をOpenAIに送信します。合計6回のAPI料金が発生します。", wraplength=620).grid(row=13, column=0, columnspan=3, sticky="w", pady=(10, 6))
            self.start_button = ttk.Button(frame, text="評価を開始する", command=self.start)
            self.start_button.grid(row=14, column=0, sticky="w", pady=6)
            self.controls.append((self.start_button, "normal"))
            self.progress = ttk.Progressbar(frame, maximum=6, mode="determinate")
            self.progress.grid(row=14, column=1, columnspan=2, sticky="ew", padx=(12, 0))
            ttk.Label(frame, textvariable=self.status, wraplength=760).grid(row=15, column=0, columnspan=3, sticky="w", pady=(6, 10))
            self.preview = ScrolledText(frame, wrap="word", height=12, state="disabled")
            self.preview.grid(row=16, column=0, columnspan=3, sticky="nsew")
            root.after(100, self.poll)

        def entry(self, frame, row, label, variable, browse=None, password=False):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
            widget = ttk.Entry(frame, textvariable=variable, show="*" if password else "")
            widget.grid(row=row, column=1, columnspan=1 if browse else 2, sticky="ew", pady=5)
            self.controls.append((widget, "normal"))
            if browse:
                button = ttk.Button(frame, text="参照…", command=browse)
                button.grid(row=row, column=2, padx=(8, 0), pady=5)
                self.controls.append((button, "normal"))

        def select_input(self):
            path = filedialog.askopenfilename(parent=self.root, title="小説を選択", filetypes=[("テキスト", "*.txt"), ("すべて", "*")])
            if path:
                self.input_path.set(path)
                source = Path(path)
                self.output_path.set(str(source.with_name(source.stem + "_批評.txt")))

        def select_output(self):
            current = Path(self.output_path.get()) if self.output_path.get() else None
            path = filedialog.asksaveasfilename(parent=self.root, title="結果の保存先", defaultextension=".txt", initialfile=current.name if current else "批評.txt", initialdir=str(current.parent) if current else None, filetypes=[("テキスト", "*.txt")])
            if path:
                self.output_path.set(path)

        def start(self):
            if self.running:
                return
            try:
                if not self.input_path.get().strip() or not self.output_path.get().strip():
                    raise ValueError("入力ファイルと保存先を指定してください。")
                try:
                    tokens = int(self.tokens.get())
                    timeout = int(self.timeout.get())
                    min_chars = int(self.min_chars.get()) if self.min_chars.get().strip() else None
                    max_chars = int(self.max_chars.get()) if self.max_chars.get().strip() else None
                except ValueError:
                    raise ValueError("トークン数・タイムアウト・文字数の閾値には整数を入力してください。") from None
                validate_thresholds(min_chars, max_chars)
                if not 256 <= tokens <= 32000 or not 1 <= timeout <= 3600:
                    raise ValueError("出力トークンは256～32000、タイムアウトは1～3600秒にしてください。")
                settings = Settings(Path(self.input_path.get()), Path(self.output_path.get()), self.api_key.get(), self.model.get(), self.encoding.get(), tokens, timeout, min_chars, max_chars)
            except (ValueError, CritiqueError) as exc:
                messagebox.showerror("設定を確認してください", str(exc), parent=self.root)
                return
            self.running = True
            self.completed = 0
            self.progress["value"] = 0
            self.set_preview("")
            self.status.set("入力ファイルと設定を確認中です…")
            for widget, _ in self.controls:
                widget.configure(state="disabled")
            threading.Thread(target=run_job, args=(settings, self.events), daemon=True).start()

        def set_preview(self, text):
            self.preview.configure(state="normal")
            self.preview.delete("1.0", "end")
            self.preview.insert("1.0", text)
            self.preview.configure(state="disabled")

        def poll(self):
            try:
                while True:
                    kind, text = self.events.get_nowait()
                    if kind == "progress":
                        self.progress["value"] = self.completed
                        self.completed += 1
                        self.status.set(text)
                    else:
                        self.running = False
                        for widget, state in self.controls:
                            widget.configure(state=state)
                        if kind in ("done", "skipped"):
                            self.progress["value"] = 6 if kind == "done" else 0
                            self.set_preview(text)
                            self.status.set(("評価無し・API送信なし。保存しました: " if kind == "skipped" else "保存しました: ") + self.output_path.get())
                        else:
                            self.status.set("処理を完了できませんでした。")
                            messagebox.showerror("評価エラー", text + "\n送信済みの処理には料金が発生する場合があります。", parent=self.root)
            except Empty:
                pass
            self.root.after(100, self.poll)

        def close(self):
            if self.running:
                messagebox.showinfo("処理中です", "保存を完了するため、処理が終わるまでお待ちください。", parent=self.root)
                return
            self.api_key.set("")
            self.root.destroy()

    root = tk.Tk()
    Application(root)
    root.mainloop()


if __name__ == "__main__":
    main()
