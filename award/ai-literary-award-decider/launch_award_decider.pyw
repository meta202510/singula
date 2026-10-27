"""Console-free Windows entry point for the award-decider GUI."""

try:
    import tkinter as tk
    from tkinter import messagebox
    from award_decider_gui import App

    root = tk.Tk()
    def report_callback_exception(_error_type, _error, _traceback):
        messagebox.showerror(
            "画面処理エラー",
            "画面処理中に予期しないエラーが発生しました。\n"
            "詳細確認が必要な場合は start_award_decider.bat から起動してください。",
            parent=root,
        )
    root.report_callback_exception = report_callback_exception
    App(root)
    root.mainloop()
except Exception:
    try:
        import tkinter as tk
        from tkinter import messagebox
        error_root = tk.Tk()
        error_root.withdraw()
        messagebox.showerror(
            "起動エラー",
            "受賞作品決定ツールを起動できませんでした。\n"
            "ZIP内のファイルを同じフォルダに置いてください。\n"
            "詳細確認が必要な場合は start_award_decider.bat を実行してください。",
            parent=error_root,
        )
        error_root.destroy()
    except Exception:
        pass
