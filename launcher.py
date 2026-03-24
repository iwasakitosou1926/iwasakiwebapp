# launcher.py
import os
import sys
import runpy
import traceback

def main():
    # In PyInstaller frozen mode:
    #   sys.executable = ...\NotionSyncApp\NotionSyncApp.exe
    #   exe_dir        = folder containing the exe (the inner onedir folder)
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))

    # OUTER install folder = parent of the exe folder
    # Expected layout:
    #   OUTER\
    #     app\app_main.py
    #     .env
    #     NotionSyncApp\NotionSyncApp.exe
    outer_dir = os.path.dirname(exe_dir)

    app_dir = os.path.join(outer_dir, "app")
    entry = os.path.join(app_dir, "app_main.py")

    # Make sure relative paths (and .env loading) behave from OUTER folder
    try:
        os.chdir(outer_dir)
    except Exception:
        pass

    # Never overwrite sys.path; only prepend our external app folder
    if app_dir not in sys.path:
        sys.path.insert(0, app_dir)

    # Helpful runtime debug files (optional but useful)
    try:
        with open(os.path.join(outer_dir, "sys_path_debug.txt"), "w", encoding="utf-8") as f:
            f.write(f"sys.executable = {sys.executable}\n")
            f.write(f"exe_dir = {exe_dir}\n")
            f.write(f"outer_dir = {outer_dir}\n")
            f.write(f"app_dir = {app_dir}\n")
            f.write(f"entry = {entry}\n\n")
            f.write("sys.path:\n" + "\n".join(sys.path) + "\n")
    except Exception:
        pass

    # Validate app exists
    if not os.path.exists(entry):
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "NotionSyncApp",
                "Missing app entrypoint:\n\n"
                f"{entry}\n\n"
                "Expected folder layout:\n"
                "OUTER\\app\\app_main.py\n"
                "OUTER\\.env\n"
                "OUTER\\NotionSyncApp\\NotionSyncApp.exe",
            )
        except Exception:
            pass
        return

    # Run the external app and capture crashes to a log file
    try:
        runpy.run_path(entry, run_name="__main__")
    except Exception:
        log_path = os.path.join(outer_dir, "app_crash.log")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except Exception:
            pass

        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "NotionSyncApp",
                "App crashed.\n\n"
                f"Log saved to:\n{log_path}"
            )
        except Exception:
            pass


if __name__ == "__main__":
    main()