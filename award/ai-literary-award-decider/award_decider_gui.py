"""Windows GUI for comparing existing critique-result text files."""
from __future__ import annotations
import os
from pathlib import Path
from queue import Empty, Queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from award_decider import AwardError, MODEL_FETCH_ERROR, discover, decide, estimate_tokens, fetch_models, format_output, make_output_path, save_text


class App:
    def __init__(self, root: tk.Tk):
        self.root = root; self.events = Queue(); self.running = False; self.fetching_models = False; self.files = []
        root.title("AI文学賞 受賞作品決定")
        root.geometry("900x760"); root.minsize(720, 520)
        root.columnconfigure(0, weight=1); root.rowconfigure(0, weight=1)
        frame = ttk.Frame(root, padding=16); frame.grid(sticky="nsew"); frame.columnconfigure(0, weight=1); frame.rowconfigure(5, weight=1)
        ttk.Label(frame, text="AI文学賞 受賞作品決定", font=("", 20, "bold")).grid(row=0, column=0, sticky="w")
        top = ttk.Frame(frame); top.grid(row=1, column=0, sticky="ew", pady=8); top.columnconfigure(1, weight=1)
        self.folder = tk.StringVar(); self.output_folder = tk.StringVar(); self.key = tk.StringVar(value=os.environ.get("OPENAI_API_KEY", "")); self.model = tk.StringVar(); self.mock = tk.BooleanVar(); self.status = tk.StringVar(value="批評結果フォルダを選択してください。"); self.count = tk.StringVar(value="読み込み件数: 0 / 総文字数: 0 / 推定入力トークン: 0")
        ttk.Label(top, text="批評結果フォルダ").grid(row=0, column=0, sticky="w"); ttk.Entry(top, textvariable=self.folder).grid(row=0, column=1, sticky="ew", padx=8); ttk.Button(top, text="参照…", command=self.pick_folder).grid(row=0, column=2)
        ttk.Label(top, text="出力先フォルダ").grid(row=1, column=0, sticky="w"); ttk.Entry(top, textvariable=self.output_folder).grid(row=1, column=1, sticky="ew", padx=8); ttk.Button(top, text="フォルダ選択…", command=self.pick_output_folder).grid(row=1, column=2)
        ttk.Label(top, text="APIキー").grid(row=2, column=0, sticky="w")
        self.key_entry = ttk.Entry(top, textvariable=self.key, show="*"); self.key_entry.grid(row=2, column=1, sticky="ew", padx=8)
        self.key_entry.bind("<Return>", lambda _: self.refresh_models())
        self.key_entry.bind("<FocusOut>", lambda _: self.refresh_models())
        ttk.Label(top, text="結果・ログには保存しません").grid(row=2, column=2)
        ttk.Label(top, text="AIモデル").grid(row=3, column=0, sticky="w")
        self.model_selector = ttk.Combobox(top, textvariable=self.model, values=(), state="readonly")
        self.model_selector.grid(row=3, column=1, sticky="ew", padx=8)
        self.model_button = ttk.Button(top, text="モデル一覧を取得", command=self.refresh_models); self.model_button.grid(row=3, column=2)
        ttk.Checkbutton(top, text="モックテスト（通信・課金なし）", variable=self.mock).grid(row=4, column=1, sticky="w")
        ttk.Label(frame, textvariable=self.count).grid(row=2, column=0, sticky="w")
        self.listbox = tk.Listbox(frame, height=7); self.listbox.grid(row=3, column=0, sticky="ew", pady=6)
        buttons = ttk.Frame(frame); buttons.grid(row=4, column=0, sticky="w"); ttk.Button(buttons, text="ファイル一覧を読み込む", command=self.load).grid(row=0, column=0, padx=(0,8)); self.run_button = ttk.Button(buttons, text="受賞作品を決定する", command=self.start); self.run_button.grid(row=0, column=1)
        self.progress = ttk.Progressbar(frame, mode="indeterminate"); self.progress.grid(row=4, column=0, sticky="e")
        ttk.Label(frame, textvariable=self.status, wraplength=820).grid(row=5, column=0, sticky="nw", pady=6)
        self.result = ScrolledText(frame, wrap="word", height=16, state="disabled"); self.result.grid(row=6, column=0, sticky="nsew"); frame.rowconfigure(6, weight=1)
        root.after(100, self.poll)
        root.after(200, self.refresh_models)

    def refresh_models(self):
        if self.fetching_models or self.running: return
        api_key = self.key.get().strip()
        if not api_key:
            self.apply_model_error()
            return
        self.fetching_models = True; self.model_button.configure(state="disabled"); self.status.set("利用可能なモデルを取得中です…")
        threading.Thread(target=self.model_worker, args=(api_key,), daemon=True).start()
    def model_worker(self, api_key):
        try: self.events.put(("models_loaded", fetch_models(api_key)))
        except AwardError: self.events.put(("models_error", MODEL_FETCH_ERROR))
    def apply_model_error(self):
        self.model_selector.configure(values=(MODEL_FETCH_ERROR,)); self.model.set(MODEL_FETCH_ERROR); self.status.set(MODEL_FETCH_ERROR)

    def pick_folder(self):
        value = filedialog.askdirectory(parent=self.root, title="批評結果フォルダを選択")
        if value:
            self.folder.set(value)
            if not self.output_folder.get(): self.output_folder.set(value)
            self.load()
    def pick_output_folder(self):
        value = filedialog.askdirectory(parent=self.root, title="選考結果を保存するフォルダを選択")
        if value: self.output_folder.set(value)
    def load(self):
        try:
            self.files = discover(Path(self.folder.get()))
            self.listbox.delete(0, "end")
            for item in self.files: self.listbox.insert("end", f"{item.path.name}  —  {item.title}")
            total = sum(len(f.text) for f in self.files); self.count.set(f"読み込み件数: {len(self.files)} / 総文字数: {total} / 推定入力トークン: {estimate_tokens(self.files)}")
            self.status.set("対象ファイルを確認しました。" if self.files else "対象ファイルがありません（APIは呼び出しません）。")
        except AwardError as exc: messagebox.showerror("読み込みエラー", str(exc), parent=self.root)
    def start(self):
        if self.running: return
        try:
            if not self.files: self.load()
            if not self.files: return
            if not self.output_folder.get().strip(): raise AwardError("出力先フォルダを選択してください。")
            selected_model = self.model.get().strip()
            if not self.mock and len(self.files) > 1 and (not selected_model or selected_model == MODEL_FETCH_ERROR): raise AwardError(MODEL_FETCH_ERROR)
            if not self.mock and len(self.files) > 1 and not self.key.get().strip(): raise AwardError("APIキーを入力してください。")
            if len(self.files) > 1 and estimate_tokens(self.files) > 100000: raise AwardError("入力容量が大きすぎるため停止しました。入力を削除・切り捨てずに、作品数やモデルの入力上限を確認してください。")
            model_for_run = selected_model if selected_model and selected_model != MODEL_FETCH_ERROR else "API未使用"
            output = make_output_path(Path(self.output_folder.get())); self.running = True; self.run_button.configure(state="disabled"); self.progress.start(10); self.status.set(f"選考を開始しました… 保存予定: {output.name}")
            threading.Thread(target=self.worker, args=(output, model_for_run), daemon=True).start()
        except AwardError as exc: messagebox.showerror("設定エラー", str(exc), parent=self.root)
    def worker(self, output, model):
        try:
            result, usage = decide(self.files, self.key.get(), model, self.mock.get(), lambda text: self.events.put(("status", text)))
            text = format_output(result, self.files, model, usage); save_text(output, text); self.events.put(("done", (text, output)))
        except AwardError as exc: self.events.put(("error", str(exc)))
        except OSError as exc: self.events.put(("error", f"ファイル保存に失敗しました: {exc}"))
    def poll(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "models_loaded":
                    self.fetching_models = False; self.model_button.configure(state="normal"); self.model_selector.configure(values=payload); self.model.set(payload[0]); self.status.set(f"モデルを{len(payload)}件取得しました。")
                elif kind == "models_error":
                    self.fetching_models = False; self.model_button.configure(state="normal"); self.apply_model_error()
                elif kind == "status": self.status.set(payload)
                else:
                    self.running = False; self.run_button.configure(state="normal"); self.progress.stop()
                    if kind == "done":
                        text, output = payload; self.result.configure(state="normal"); self.result.delete("1.0", "end"); self.result.insert("1.0", text); self.result.configure(state="disabled"); self.status.set(f"選考結果を保存しました: {output}")
                    else: self.status.set("選考に失敗しました。"); messagebox.showerror("選考エラー", payload, parent=self.root)
        except Empty: pass
        self.root.after(100, self.poll)


if __name__ == "__main__":
    root = tk.Tk(); App(root); root.mainloop()
