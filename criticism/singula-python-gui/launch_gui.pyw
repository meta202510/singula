"""Console-free Windows entry point for the individual critique GUI."""

try:
    from novel_critique_gui import main
    main()
except Exception:
    # Keep request text, API keys and raw exception details out of dialogs/logs.
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "起動エラー",
            "個別評価ツールを起動できませんでした。\n"
            "ZIP内のファイルを同じフォルダに置いてください。\n"
            "詳細確認が必要な場合は start_gui.bat を実行してください。",
            parent=root,
        )
        root.destroy()
    except Exception:
        pass
