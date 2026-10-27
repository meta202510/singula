"""Windows GUI for comparing existing critique-result text files."""
from __future__ import annotations
import os
from pathlib import Path
from queue import Empty, Queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from award_decider import AwardError, discover, decide, estimate_tokens, format_output, save_text

MODEL_CHOICES = (
    "gpt-4.1-mini",
    "gpt-4.1",
    "gpt-4.1-nano",
    "gpt-5",
    "gpt-5-mini",
    "gpt-5-nano",
)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root; self.events = Queue(); self.running = False; self.files = []
        root.title("AI文学賞 受賞作品決定")
        root.geometry("900x760"); root.minsize(720, 520)
        root.columnconfigure(0, weight=1); root.rowconfigure(0, weight=1)
        frame = ttk.Frame(root, padding=16); frame.grid(sticky="nsew"); frame.columnconfigure(0, weight=1); frame.rowconfigure(5, weight=1)
        ttk.Label(frame, text="AI文学賞 受賞作品決定", font=("", 20, "bold")).grid(row=0, column=0, sticky="w")
        top = ttk.Frame(frame); top.grid(row=1, column=0, sticky="ew", pady=8); top.columnconfigure(1, weight=1)
        self.folder = tk.StringVar(); self.output = tk.StringVar(); self.key = tk.StringVar(value=os.environ.get("OPENAI_API_KEY", "")); self.model = tk.StringVar(value="gpt-4.1-mini"); self.mock = tk.BooleanVar(); self.status = tk.StringVar(value="批評結果フォルダを選択してください。"); self.count = tk.StringVar(value="読み込み件数: 0 / 総文字数: 0 / 推定入力トークン: 0")
        ttk.Label(top, text="批評結果フォルダ").grid(row=0, column=0, sticky="w"); ttk.Entry(top, textvariable=self.folder).grid(row=0, column=1, sticky="ew", padx=8); ttk.Button(top, text="参照…", command=self.pick_folder).grid(row=0, column=2)
        ttk.Label(top, text="出力先").grid(row=1, column=0, sticky="w"); ttk.Entry(top, textvariable=self.output).grid(row=1, column=1, sticky="ew", padx=8); ttk.Button(top, text="保存先…", command=self.pick_output).grid(row=1, column=2)
        ttk.Label(top, text="APIキー").grid(row=2, column=0, sticky="w"); ttk.Entry(top, textvariable=self.key, show="*").grid(row=2, column=1, sticky="ew", padx=8); ttk.Label(top, text="結果・ログには保存しません").grid(row=2, column=2)
        ttk.Label(top, text="AIモデル（直接入力も可）").grid(row=3, column=0, sticky="w")
        self.model_selector = ttk.Combobox(top, textvariable=self.model, values=MODEL_CHOICES, state="normal")
        self.model_selector.grid(row=3, column=1, sticky="ew", padx=8)
        ttk.Checkbutton(top, text="モックテスト（通信・課金なし）", variable=self.mock).grid(row=4, column=1, sticky="w")
        ttk.Label(frame, textvariable=self.count).grid(row=2, column=0, sticky="w")
        self.listbox = tk.Listbox(frame, height=7); self.listbox.grid(row=3, column=0, sticky="ew", pady=6)
        buttons = ttk.Frame(frame); buttons.grid(row=4, column=0, sticky="w"); ttk.Button(buttons, text="ファイル一覧を読み込む", command=self.load).grid(row=0, column=0, padx=(0,8)); self.run_button = ttk.Button(buttons, text="受賞作品を決定する", command=self.start); self.run_button.grid(row=0, column=1)
        self.progress = ttk.Progressbar(frame, mode="indeterminate"); self.progress.grid(row=4, column=0, sticky="e")
        ttk.Label(frame, textvariable=self.status, wraplength=820).grid(row=5, column=0, sticky="nw", pady=6)
        self.result = ScrolledText(frame, wrap="word", height=16, state="disabled"); self.result.grid(row=6, column=0, sticky="nsew"); frame.rowconfigure(6, weight=1)
        root.after(100, self.poll)

    def pick_folder(self):
        value = filedialog.askdirectory(parent=self.root, title="批評結果フォルダを選択")
        if value: self.folder.set(value); self.output.set(str(Path(value) / "受賞作品決定結果.txt")); self.load()
    def pick_output(self):
        value = filedialog.asksaveasfilename(parent=self.root, title="選考結果の保存先", defaultextension=".txt", filetypes=[("テキスト", "*.txt")])
        if value: self.output.set(value)
    def load(self):
        try:
            output = Path(self.output.get()) if self.output.get() else None; self.files = discover(Path(self.folder.get()), output)
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
            if not self.output.get().strip(): raise AwardError("出力先を指定してください。")
            if not self.model.get().strip(): raise AwardError("AIモデルを選択または入力してください。")
            if not self.mock and len(self.files) > 1 and not self.key.get().strip(): raise AwardError("APIキーを入力してください。")
            if len(self.files) > 1 and estimate_tokens(self.files) > 100000: raise AwardError("入力容量が大きすぎるため停止しました。入力を削除・切り捨てずに、作品数やモデルの入力上限を確認してください。")
            output = Path(self.output.get()); self.running = True; self.run_button.configure(state="disabled"); self.progress.start(10); self.status.set("選考を開始しました…")
            threading.Thread(target=self.worker, args=(output,), daemon=True).start()
        except AwardError as exc: messagebox.showerror("設定エラー", str(exc), parent=self.root)
    def worker(self, output):
        try:
            result, usage = decide(self.files, self.key.get(), self.model.get(), self.mock.get(), lambda text: self.events.put(("status", text)))
            text = format_output(result, self.files, self.model.get(), usage); save_text(output, text); self.events.put(("done", text))
        except AwardError as exc: self.events.put(("error", str(exc)))
        except OSError as exc: self.events.put(("error", f"ファイル保存に失敗しました: {exc}"))
    def poll(self):
        try:
            while True:
                kind, text = self.events.get_nowait()
                if kind == "status": self.status.set(text)
                else:
                    self.running = False; self.run_button.configure(state="normal"); self.progress.stop()
                    if kind == "done": self.result.configure(state="normal"); self.result.delete("1.0", "end"); self.result.insert("1.0", text); self.result.configure(state="disabled"); self.status.set("選考結果を保存しました。")
                    else: self.status.set("選考に失敗しました。"); messagebox.showerror("選考エラー", text, parent=self.root)
        except Empty: pass
        self.root.after(100, self.poll)


if __name__ == "__main__":
    root = tk.Tk(); App(root); root.mainloop()
