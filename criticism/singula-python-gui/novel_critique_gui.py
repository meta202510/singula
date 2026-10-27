"""Local desktop interface. Run: python novel_critique_gui.py"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
import os
from pathlib import Path
from queue import Empty, Queue
import threading

from novel_critique import (
    DEFAULT_MODEL, MODEL_FETCH_ERROR, CritiqueError, OpenAIReviewer, evaluate, fetch_models, format_result, save_result,
    threshold_result, validate_thresholds,
)
from mock_review import MARKER, MockOptions, MockStopped, evaluate_mock


def model_loaded_message(models) -> str:
    return f"モデルを{len(models)}件取得しました。"

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
    mock: bool = False
    mock_resume: bool = False
    mock_stop_after: int | None = None
    mock_fail_at: int | None = None


def run_job(settings: Settings, events: Queue) -> None:
    """Worker: no Tk calls; all user-visible changes travel through a queue."""
    try:
        mock_options = MockOptions(settings.mock_resume, settings.mock_stop_after, settings.mock_fail_at)
        if settings.mock:
            mock_options.validate()
        elif settings.mock_resume or settings.mock_stop_after is not None or settings.mock_fail_at is not None:
            raise CritiqueError("モック用の設定にはモックモードを選択してください。")
        if settings.input_path.resolve() == settings.output_path.resolve():
            raise CritiqueError("入力ファイルと異なる保存先を選択してください。")
        if os.path.lexists(settings.output_path):
            raise CritiqueError("保存先が既に存在します。別の名前を選択してください。")
        if not settings.output_path.parent.is_dir():
            raise CritiqueError("保存先のフォルダーが存在しません。")
        if not settings.mock and (not settings.model.strip() or settings.model == MODEL_FETCH_ERROR):
            raise CritiqueError(MODEL_FETCH_ERROR)
        if settings.encoding not in ("utf-8-sig", "cp932", "shift_jis"):
            raise CritiqueError("入力文字コードを選択してください。")
        if not 256 <= settings.max_output_tokens <= 32000 or not 0 < settings.timeout <= 3600:
            raise CritiqueError("出力トークンは256～32000、タイムアウトは1～3600秒にしてください。")
        novel = settings.input_path.read_text(encoding=settings.encoding)
        skipped = threshold_result(settings.input_path.name, novel, settings.min_chars, settings.max_chars)
        if skipped is not None:
            if settings.mock:
                skipped += MARKER + "\n"
            save_result(settings.output_path, skipped)
            events.put(("skipped", skipped))
            return
        if not novel.strip():
            raise CritiqueError("入力ファイルが空です。")
        if settings.mock:
            reviews, overall = evaluate_mock(novel, settings.input_path.name, settings.output_path, mock_options, lambda text: events.put(("progress", text)))
        else:
            if not settings.api_key.strip():
                raise CritiqueError("APIキーを入力するか、環境変数 OPENAI_API_KEY を設定してください。")
            reviewer = OpenAIReviewer(
                settings.api_key.strip(), settings.model.strip(), settings.max_output_tokens, settings.timeout,
            )
            reviews, overall = evaluate(novel, reviewer.review, lambda text: events.put(("progress", text)))
        result = format_result(settings.input_path.name, reviews, overall)
        save_result(settings.output_path, result)
        events.put(("done", result))
    except MockStopped as exc:
        events.put(("paused", str(exc)))
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
            self.fetching_models = False
            self.events = Queue()
            self.controls = []
            self.completed = 0
            root.title("小説批評 — 5人の審査員と総評")
            height = min(780, max(360, root.winfo_screenheight() - 140))
            root.geometry(f"900x{height}")
            root.minsize(760, 360)
            root.protocol("WM_DELETE_WINDOW", self.close)
            root.columnconfigure(0, weight=1)
            root.rowconfigure(0, weight=1)
            self.canvas = tk.Canvas(root, highlightthickness=0)
            self.canvas.grid(row=0, column=0, sticky="nsew")
            scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")
            self.canvas.configure(yscrollcommand=scrollbar.set)
            frame = ttk.Frame(self.canvas, padding=18)
            self.content = frame
            self.content_window = self.canvas.create_window((0, 0), window=frame, anchor="nw")
            frame.bind("<Configure>", self.resize_content)
            self.canvas.bind("<Configure>", self.resize_content)
            # Toplevel bindings also receive events from descendant controls.
            root.bind("<MouseWheel>", self.scroll_wheel, add="+")
            root.bind("<Button-4>", self.scroll_wheel, add="+")
            root.bind("<Button-5>", self.scroll_wheel, add="+")
            frame.columnconfigure(1, weight=1)
            frame.rowconfigure(16, weight=1)
            ttk.Label(frame, text="小説批評", font=("", 20, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
            ttk.Label(frame, text="5人が本文を独立して批評し、その批評だけから総評を作成します。").grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 16))
            self.input_path = tk.StringVar()
            self.output_path = tk.StringVar()
            self.api_key = tk.StringVar(value=os.environ.get("OPENAI_API_KEY", ""))
            self.model = tk.StringVar()
            self.encoding = tk.StringVar(value="utf-8-sig")
            self.tokens = tk.StringVar(value="4000")
            self.timeout = tk.StringVar(value="180")
            self.min_chars = tk.StringVar()
            self.max_chars = tk.StringVar()
            self.mock = tk.BooleanVar(value=False)
            self.mock_resume = tk.BooleanVar(value=False)
            self.mock_stop = tk.StringVar()
            self.mock_fail = tk.StringVar()
            self.mode_notice = tk.StringVar(value="本番：範囲内の本文をOpenAIに送信します。6回分のAPI料金が発生します。")
            self.status = tk.StringVar(value="入力ファイルと保存先を選択してください。")
            self.entry(frame, 2, "小説ファイル", self.input_path, self.select_input)
            self.entry(frame, 3, "結果の保存先", self.output_path, self.select_output)
            self.api_key_entry = self.entry(frame, 4, "APIキー", self.api_key, password=True)
            self.api_key_entry.bind("<Return>", lambda _: self.refresh_models())
            self.api_key_entry.bind("<FocusOut>", lambda _: self.refresh_models())
            ttk.Label(frame, text="キーはこの画面でのみ保持し、設定ファイルには保存しません。").grid(row=5, column=1, columnspan=2, sticky="w")
            ttk.Label(frame, text="AIモデル").grid(row=6, column=0, sticky="w", padx=(0, 12), pady=5)
            self.model_selector = ttk.Combobox(frame, textvariable=self.model, values=(), state="readonly")
            self.model_selector.grid(row=6, column=1, sticky="ew", pady=5)
            self.model_button = ttk.Button(frame, text="モデル一覧を取得", command=self.refresh_models)
            self.model_button.grid(row=6, column=2, padx=(8, 0), pady=5)
            self.controls.extend([(self.model_selector, "readonly"), (self.model_button, "normal")])
            ttk.Label(frame, text="入力文字コード").grid(row=7, column=0, sticky="w", padx=(0, 12), pady=5)
            encoding = ttk.Combobox(frame, textvariable=self.encoding, values=("utf-8-sig", "cp932", "shift_jis"), state="readonly")
            encoding.grid(row=7, column=1, columnspan=2, sticky="ew", pady=5)
            self.controls.append((encoding, "readonly"))
            self.entry(frame, 8, "最大出力トークン／評価", self.tokens)
            self.entry(frame, 9, "通信タイムアウト（秒）", self.timeout)
            self.entry(frame, 10, "入力文字数の下限", self.min_chars)
            self.entry(frame, 11, "入力文字数の上限", self.max_chars)
            ttk.Label(frame, text="閾値は空欄で制限なし。範囲外はAPIを呼ばず「評価無し」を保存します。").grid(row=12, column=0, columnspan=3, sticky="w")
            mode = ttk.Frame(frame)
            mode.grid(row=13, column=0, columnspan=3, sticky="ew", pady=6)
            mock_button = ttk.Checkbutton(mode, text="モックテスト（通信・課金なし）", variable=self.mock, command=self.toggle_mock)
            mock_button.grid(row=0, column=0, columnspan=2, sticky="w")
            resume_button = ttk.Checkbutton(mode, text="モック途中データから再開", variable=self.mock_resume)
            resume_button.grid(row=0, column=2, columnspan=2, sticky="w")
            self.controls.extend([(mock_button, "normal"), (resume_button, "normal")])
            for column, label, variable in [(0, "停止する回（完了後）", self.mock_stop), (2, "エラーを起こす回", self.mock_fail)]:
                ttk.Label(mode, text=label).grid(row=1, column=column, sticky="w")
                combo = ttk.Combobox(mode, textvariable=variable, values=("", "1", "2", "3", "4", "5", "6"), width=4, state="readonly")
                combo.grid(row=1, column=column + 1, padx=8)
                self.controls.append((combo, "readonly"))
            ttk.Label(mode, textvariable=self.mode_notice, wraplength=740).grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))
            self.start_button = ttk.Button(frame, text="評価を開始する", command=self.start)
            self.start_button.grid(row=14, column=0, sticky="w", pady=6)
            self.controls.append((self.start_button, "normal"))
            self.progress = ttk.Progressbar(frame, maximum=6, mode="determinate")
            self.progress.grid(row=14, column=1, columnspan=2, sticky="ew", padx=(12, 0))
            ttk.Label(frame, textvariable=self.status, wraplength=760).grid(row=15, column=0, columnspan=3, sticky="w", pady=(6, 10))
            self.preview = ScrolledText(frame, wrap="word", height=6, state="disabled")
            self.preview.grid(row=16, column=0, columnspan=3, sticky="nsew")
            root.after(100, self.poll)
            root.after(200, self.refresh_models)

        def refresh_models(self):
            if self.fetching_models or self.running:
                return
            api_key = self.api_key.get().strip()
            if not api_key:
                self.apply_model_error()
                return
            self.fetching_models = True
            self.model_button.configure(state="disabled")
            self.status.set("利用可能なモデルを取得中です…")
            threading.Thread(target=self.model_worker, args=(api_key,), daemon=True).start()

        def model_worker(self, api_key):
            try:
                self.events.put(("models_loaded", fetch_models(api_key)))
            except CritiqueError:
                self.events.put(("models_error", MODEL_FETCH_ERROR))

        def apply_model_error(self):
            self.model_selector.configure(values=(MODEL_FETCH_ERROR,))
            self.model.set(MODEL_FETCH_ERROR)
            self.status.set(MODEL_FETCH_ERROR)

        def resize_content(self, event=None):
            self.canvas.itemconfigure(
                self.content_window,
                width=self.canvas.winfo_width(),
                height=max(self.content.winfo_reqheight(), self.canvas.winfo_height()),
            )
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def scroll_wheel(self, event):
            # Let the result text and dropdowns keep their native scrolling.
            if isinstance(event.widget, (tk.Text, ttk.Combobox)):
                return
            if event.widget.winfo_toplevel() != self.root:
                return
            if self.content.winfo_reqheight() <= self.canvas.winfo_height():
                return
            if getattr(event, "num", None) == 4:
                units = -1
            elif getattr(event, "num", None) == 5:
                units = 1
            else:
                delta = getattr(event, "delta", 0)
                if not delta:
                    return
                units = -int(delta / 120) if abs(delta) >= 120 else (-1 if delta > 0 else 1)
            self.canvas.yview_scroll(units, "units")
            return "break"

        def entry(self, frame, row, label, variable, browse=None, password=False):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
            widget = ttk.Entry(frame, textvariable=variable, show="*" if password else "")
            widget.grid(row=row, column=1, columnspan=1 if browse else 2, sticky="ew", pady=5)
            self.controls.append((widget, "normal"))
            if browse:
                button = ttk.Button(frame, text="参照…", command=browse)
                button.grid(row=row, column=2, padx=(8, 0), pady=5)
                self.controls.append((button, "normal"))
            return widget

        def select_input(self):
            path = filedialog.askopenfilename(parent=self.root, title="小説を選択", filetypes=[("テキスト", "*.txt"), ("すべて", "*")])
            if path:
                self.input_path.set(path)
                source = Path(path)
                self.output_path.set(str(source.with_name(source.stem + ("_MOCK_批評.txt" if self.mock.get() else "_批評.txt"))))

        def toggle_mock(self):
            self.mode_notice.set("モック：通信・課金なし。出力はテスト専用です。" if self.mock.get() else "本番：範囲内の本文をOpenAIに送信します。6回分のAPI料金が発生します。")
            if not self.mock.get():
                self.mock_resume.set(False)
                self.mock_stop.set("")
                self.mock_fail.set("")
            if self.input_path.get():
                source = Path(self.input_path.get())
                old = source.with_name(source.stem + ("_批評.txt" if self.mock.get() else "_MOCK_批評.txt"))
                if self.output_path.get() == str(old):
                    self.output_path.set(str(source.with_name(source.stem + ("_MOCK_批評.txt" if self.mock.get() else "_批評.txt"))))

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
                selected_model = self.model.get().strip()
                if not self.mock.get() and (not selected_model or selected_model == MODEL_FETCH_ERROR):
                    raise CritiqueError(MODEL_FETCH_ERROR)
                model_for_run = selected_model if selected_model and selected_model != MODEL_FETCH_ERROR else "API未使用"
                settings = Settings(Path(self.input_path.get()), Path(self.output_path.get()), self.api_key.get(), model_for_run, self.encoding.get(), tokens, timeout, min_chars, max_chars)
                settings = replace(settings, mock=self.mock.get(), mock_resume=self.mock_resume.get(), mock_stop_after=int(self.mock_stop.get()) if self.mock_stop.get() else None, mock_fail_at=int(self.mock_fail.get()) if self.mock_fail.get() else None)
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
                    if kind == "models_loaded":
                        self.fetching_models = False
                        self.model_button.configure(state="normal")
                        self.model_selector.configure(values=text)
                        self.model.set(text[0])
                        self.status.set(model_loaded_message(text))
                    elif kind == "models_error":
                        self.fetching_models = False
                        self.model_button.configure(state="normal")
                        self.apply_model_error()
                    elif kind == "progress":
                        self.progress["value"] = self.completed
                        self.completed += 1
                        self.status.set(text)
                    else:
                        self.running = False
                        for widget, state in self.controls:
                            widget.configure(state=state)
                        if kind == "paused":
                            self.status.set(text)
                            self.progress["value"] = int(self.mock_stop.get()) if self.mock_stop.get() else 0
                        elif kind in ("done", "skipped"):
                            self.progress["value"] = 6 if kind == "done" else 0
                            self.set_preview(text)
                            self.status.set(("[モック] " if self.mock.get() else "") + ("評価無し・API送信なし。保存しました: " if kind == "skipped" else "保存しました: ") + self.output_path.get())
                        else:
                            self.status.set("処理を完了できませんでした。")
                            messagebox.showerror("評価エラー", text + ("\nモックのため通信・課金はありません。" if self.mock.get() else "\n送信済みの処理には料金が発生する場合があります。"), parent=self.root)
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
    def report_callback_exception(_error_type, _error, _traceback):
        messagebox.showerror(
            "画面処理エラー",
            "画面処理中に予期しないエラーが発生しました。\n"
            "詳細確認が必要な場合は start_gui.bat から起動してください。",
            parent=root,
        )
    root.report_callback_exception = report_callback_exception
    Application(root)
    root.mainloop()


if __name__ == "__main__":
    main()
