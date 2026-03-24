# app_main.py
"""
Main entrypoint for the modular Painting Team App.

Modules:
- logic.py                   ← backend helpers (Notion, Excel, rich text, etc.)
- workflow_manager.py        ← Workflow Manager feature (manual sync)
- daily_workflow_generator.py← Daily Workflow Generator feature (auto plan)
- layout_themes.py           ← themes, DPI, gradient helpers

This file:
- Creates the main Tk window (theme + DPI-aware).
- Builds shared UI (status bar, progress bar, log console).
- Wires those into logic/workflow_manager/daily_workflow_generator.
- Provides a simple main menu to choose features.
"""

from __future__ import annotations
from layout_themes import apply_ui_scale
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from datetime import datetime, date
import threading

import logic
import workflow_manager as wf
import daily_workflow_generator as daily

# =====================================================
# DIALOG DIRECTORY NORMALIZATION HELPER
# =====================================================

def _normalize_dialog_dir(p: str | None) -> str | None:
    """Best-effort normalize for Tk filedialog initialdir.

    On Windows, network paths may be provided like //SV-04/share/.. from mac.
    Tk on Windows prefers UNC (\\SV-04\share\..). We try multiple variants
    and return the first that exists. If none exist, return the most reasonable
    normalized variant (so the intent is preserved).
    """
    if not p:
        return None
    try:
        s = str(p).strip()
    except Exception:
        return None
    if not s:
        return None

    candidates: list[str] = []
    # original
    candidates.append(s)
    # slash to backslash
    candidates.append(s.replace("/", "\\"))
    # UNC: //SERVER\share -> \\SERVER\share
    if s.startswith("//"):
        candidates.append("\\\\" + s[2:].replace("/", "\\"))
    if s.startswith("\\\\"):
        candidates.append(s)

    # normpath each
    normed: list[str] = []
    for c in candidates:
        try:
            normed.append(os.path.normpath(c))
        except Exception:
            normed.append(c)

    # prefer an existing directory
    for c in normed:
        try:
            if os.path.isdir(c):
                return c
        except Exception:
            continue

    # last resort: return first normalized candidate
    return normed[0] if normed else None

# =====================================================
# EXCEL FOLDER PREFERENCE (compat wrapper)
# Some older builds only have set_user_selected_excel_folder() and FOLDER_PATH.
# =====================================================

def _get_excel_folder_pref() -> str | None:
    """Return the preferred Excel folder.

    If logic.get_user_selected_excel_folder exists, use it.
    Otherwise, fall back to logic.FOLDER_PATH.
    """
    try:
        getter = getattr(logic, "get_user_selected_excel_folder", None)
        if callable(getter):
            v = getter()
            if v:
                return str(v)
    except Exception:
        pass
    try:
        v = getattr(logic, "FOLDER_PATH", None)
        return str(v) if v else None
    except Exception:
        return None


def _set_excel_folder_pref(p: str) -> bool:
    """Persist the preferred Excel folder if supported; always update logic.FOLDER_PATH."""
    ok = True
    try:
        setter = getattr(logic, "set_user_selected_excel_folder", None)
        if callable(setter):
            ok = bool(setter(p))
    except Exception:
        ok = False
    try:
        logic.FOLDER_PATH = p
    except Exception:
        pass
    return ok
# =====================================================
# RUNTIME VERSION (prefer version.txt written by updater)
# =====================================================
from layout_themes import (
    THEMES,
    CURRENT_THEME,
    UI_FONT_FAMILY,
    APP_VERSION,
    ensure_dpi_awareness_and_scaling,
    apply_theme,
    toggle_theme,
)

try:
    APP_VERSION_RUNTIME = (logic.read_local_version_from_install() or APP_VERSION).strip()
except Exception:
    APP_VERSION_RUNTIME = APP_VERSION

# =====================================================
# GLOBALS (FRAME / WIDGET REFS)
# =====================================================


root: tk.Tk | None = None
main_frame: ttk.Frame | None = None
status_label: ttk.Label | None = None
progress: ttk.Progressbar | None = None
log_console: tk.Text | None = None

# =====================================================
# UI ZOOM (dynamic font scaling)
# =====================================================
UI_ZOOM_FACTOR: float = 1.56  # default zoom used across screens (matches slider default)

_current_screen_cb = None  # type: ignore
_ui_zoom_var: tk.DoubleVar | None = None

def set_current_screen(cb):
    """Remember the current screen render function so zoom/theme can rebuild it."""
    global _current_screen_cb
    _current_screen_cb = cb

def UFont(size: int, *styles):
    """Return a font tuple scaled by UI_ZOOM_FACTOR."""
    try:
        s = int(round(float(size) * float(UI_ZOOM_FACTOR)))
    except Exception:
        s = size
    # Keep a sane minimum so text never becomes unreadable
    s = max(8, s)
    return (UI_FONT_FAMILY, s, *styles)


# =====================================================
# LOGGING
# =====================================================

def ui_log(msg: str):
    """
    Central logging function:
    - Prints to console (for debug).
    - Appends to the Tk log text widget.
    """
    print(msg)
    if log_console and log_console.winfo_exists():
        log_console.insert("end", msg + "\n")
        log_console.see("end")

try:
    ui_log(f"[Version] APP_VERSION_RUNTIME = {APP_VERSION_RUNTIME}")
except Exception:
    pass


# =====================================================
# LAN UPDATE CHECK (in-app notification)
# =====================================================

_update_last_info: dict | None = None


def _update_notify(info: dict):
    """Show a friendly in-app prompt if an update is available."""
    global _update_last_info
    _update_last_info = info

    if not info.get("available"):
        return

    remote = info.get("remote", "")
    notes = info.get("notes", "")
    zip_path = info.get("zip_path", "")

    msg = (
        f"アップデートがあります。\n\n"
        f"現在: {info.get('current')}\n"
        f"最新: {remote}\n\n"
        f"内容: {notes or '（説明なし）'}\n\n"
        "今すぐ更新してアプリを再起動しますか？"
    )

    if messagebox.askyesno("Update", msg):
        ok, err = logic.apply_lan_update(zip_path, remote)
        if ok:
            # Give the updater a moment to start, then exit.
            try:
                if root is not None:
                    root.after(300, lambda: root.destroy())
            except Exception:
                pass
        else:
            messagebox.showerror("Update", f"更新に失敗しました:\n{err}")


def check_updates_silent():
    """Check LAN manifest in a background thread; notify inside the app if newer."""

    def worker():
        try:
            info = logic.check_update_available(APP_VERSION_RUNTIME)
        except Exception as e:
            info = {"available": False, "error": str(e)}

        def on_ui():
            # Optional: status bar hint (non-blocking)
            if info.get("available"):
                set_status("🔔 アップデートあり（更新できます）")
                try:
                    ui_log(f"[Update] AVAILABLE current={info.get('current')} remote={info.get('remote')}")
                    ui_log(f"[Update] manifest_path = {info.get('manifest_path')}")
                except Exception:
                    pass
                _update_notify(info)
            else:
                try:
                    # Always show details so we know what was checked
                    if info.get("error"):
                        ui_log(f"[Update] ERROR: {info.get('error')}")
                    else:
                        ui_log("[Update] No updates")
                        mp = info.get("manifest_path")
                        if mp:
                            ui_log(f"[Update] manifest_path = {mp}")
                        ui_log(f"[Update] current = {info.get('current')}")
                        ui_log(f"[Update] remote  = {info.get('remote')}")
                except Exception:
                    pass

        if root is not None:
            root.after(0, on_ui)

    threading.Thread(target=worker, daemon=True).start()


# =====================================================
# MAIN MENU & NAVIGATION
# =====================================================

def clear_main_frame():
    if main_frame and main_frame.winfo_exists():
        for w in main_frame.winfo_children():
            w.destroy()


def set_status(text: str, color: str | None = None):
    if not status_label or not status_label.winfo_exists():
        return
    if color is None:
        color = THEMES[CURRENT_THEME]["text"]
    status_label.config(text=text, foreground=color)
    status_label.update_idletasks()


def set_progress(value: int | None = None, mode: str = "determinate"):
    if not progress or not progress.winfo_exists():
        return
    if mode == "indeterminate":
        progress.config(mode="indeterminate")
        progress.start(12)
    else:
        progress.stop()
        progress.config(mode="determinate")
        if value is not None:
            progress["value"] = value


def show_main_menu():
    clear_main_frame()
    set_current_screen(show_main_menu)

    outer = ttk.Frame(main_frame)
    outer.pack(fill="both", expand=True, padx=16, pady=16)

    header = ttk.Frame(outer, style="Card.TFrame")
    header.pack(fill="x", pady=(0, 16), padx=4)

    # ===============================
    # UI ZOOM (buttons)
    # ===============================
    global UI_ZOOM_FACTOR

    zoom_frame = ttk.Frame(header)
    zoom_frame.pack(side="right", padx=(10, 16))

    ttk.Label(zoom_frame, text="UI Zoom", font=UFont(9)).pack(anchor="center")

    controls = ttk.Frame(zoom_frame)
    controls.pack(anchor="center", pady=(2, 0))

    ZOOM_MIN = 1.0
    ZOOM_MAX = 1.75
    ZOOM_STEP = 0.05  # ~5% per click

    zoom_value_lbl = ttk.Label(
        controls,
        text=f"{int(round(UI_ZOOM_FACTOR * 100))}%",
        font=UFont(9, "bold"),
        width=5,
        anchor="center",
    )

    def _apply_zoom(new_val: float):
        global UI_ZOOM_FACTOR
        UI_ZOOM_FACTOR = max(ZOOM_MIN, min(ZOOM_MAX, float(new_val)))
        # Apply Tk scaling (helps some widgets), then rebuild to update explicit fonts
        if root is not None:
            apply_ui_scale(root, UI_ZOOM_FACTOR)
        zoom_value_lbl.config(text=f"{int(round(UI_ZOOM_FACTOR * 100))}%")
        if _current_screen_cb:
            _current_screen_cb()

    def _zoom_in():
        _apply_zoom(UI_ZOOM_FACTOR + ZOOM_STEP)

    def _zoom_out():
        _apply_zoom(UI_ZOOM_FACTOR - ZOOM_STEP)

    btn_minus = ttk.Button(controls, text="-", width=3, style="Menu.TButton", command=_zoom_out)
    btn_plus = ttk.Button(controls, text="+", width=3, style="Menu.TButton", command=_zoom_in)

    btn_minus.pack(side="left")
    zoom_value_lbl.pack(side="left", padx=6)
    btn_plus.pack(side="left")

    title = ttk.Label(
        header,
        text="塗装チームアプリ",
        font=UFont(18, "bold"),
        anchor="w",
    )
    title.pack(anchor="w", padx=12, pady=(8, 0))

    subtitle = ttk.Label(
        header,
        text="現場向けワークフロー / 在庫サポート用のメイン画面です。",
        font=UFont(10),
        wraplength=520,
        justify="left",
    )
    subtitle.pack(anchor="w", padx=12, pady=(2, 4))

    # Excel folder picker (shared by Workflow + Daily)
    folder_row = ttk.Frame(header)
    folder_row.pack(fill="x", padx=12, pady=(4, 0))

    excel_folder_lbl = ttk.Label(
        folder_row,
        text=f"📂 Excelフォルダ: {_get_excel_folder_pref() or ''}",
        font=UFont(9),
        wraplength=720,
        justify="left",
    )
    excel_folder_lbl.pack(side="left", anchor="w")

    def _choose_excel_folder():
        # Use normalized dialog dir for initialdir
        init_dir = _normalize_dialog_dir(_get_excel_folder_pref())
        if not init_dir:
            init_dir = os.path.expanduser("~")
        p = filedialog.askdirectory(title="Excelフォルダを選択", initialdir=init_dir)
        if not p:
            return
        ok = _set_excel_folder_pref(p)
        if not ok:
            messagebox.showerror("フォルダ", "選択したフォルダが無効です。もう一度選択してください。")
            return
        excel_folder_lbl.config(text=f"📂 Excelフォルダ: {_get_excel_folder_pref() or ''}")
        set_status(f"Excelフォルダを設定しました: {_get_excel_folder_pref() or ''}")
        ui_log(f"[PATH] Excel folder set to: {_get_excel_folder_pref() or ''}")

    ttk.Button(
        folder_row,
        text="フォルダ変更",
        style="Menu.TButton",
        command=_choose_excel_folder,
        width=10,
    ).pack(side="right")

    ver = ttk.Label(
        header,
        text=f"Version {APP_VERSION_RUNTIME}",
        font=UFont(9),
        foreground=THEMES[CURRENT_THEME]["accent"],
    )
    ver.pack(anchor="w", padx=12, pady=(0, 8))

    # Cards container (card-like layout)
    cards = ttk.Frame(outer)
    cards.pack(fill="x", pady=8)

    wf_card = ttk.Frame(cards, style="Card.TFrame")
    daily_card = ttk.Frame(cards, style="Card.TFrame")
    wf_card.pack(side="left", expand=True, fill="both", padx=(0, 8))
    daily_card.pack(side="left", expand=True, fill="both", padx=(8, 0))

    # Workflow Manager card
    ttk.Label(
        wf_card,
        text="🧩 作業フロー管理",
        font=UFont(12, "bold")
    ).pack(anchor="w", pady=(8, 2), padx=12)
    ttk.Label(
        wf_card,
        text="Excelとカレンダーを使って手動で同期します。",
        font=UFont(10),
        wraplength=260,
        justify="left"
    ).pack(anchor="w", padx=12)
    ttk.Button(
        wf_card,
        text="開く",
        style="Accent.TButton",
        command=show_workflow_manager_home,
        width=18
    ).pack(pady=(8, 12), anchor="e", padx=12)

    # Daily Generator card
    ttk.Label(
        daily_card,
        text="📅 Daily Workflow Generator",
        font=UFont(12, "bold")
    ).pack(anchor="w", pady=(8, 2), padx=12)
    ttk.Label(
        daily_card,
        text="日別のExcelから自動で『作業内容』を作成します。",
        font=UFont(10),
        wraplength=260,
        justify="left"
    ).pack(anchor="w", padx=12)
    ttk.Button(
        daily_card,
        text="開く",
        style="Accent.TButton",
        command=show_daily_generator_home,
        width=18
    ).pack(pady=(8, 12), anchor="e", padx=12)

    ttk.Label(
        outer,
        text="使いたい機能のカードを選んでください。",
        font=UFont(10)
    ).pack(pady=(12, 0))

    set_status("メインメニューを表示しました。")

    # Check for LAN updates shortly after showing the main menu
    if root is not None:
        root.after(800, check_updates_silent)


# =====================================================
# WORKFLOW MANAGER SCREEN (manual sync)
# =====================================================

_selected_calendar_for_wf: dict | None = None   # {"id": ..., "title": ..., "date": ...}

def show_workflow_manager_home():
    clear_main_frame()
    set_current_screen(show_workflow_manager_home)

    wf.wf_progress["products_loaded"] = False  # reset flag

    outer = ttk.Frame(main_frame)
    outer.pack(fill="both", expand=True, padx=16, pady=16)

    # Top row
    top = ttk.Frame(outer)
    top.pack(fill="x", pady=(0, 8))

    ttk.Button(
        top,
        text="← メニューへ",
        style="Menu.TButton",
        command=show_main_menu
    ).pack(side="left")

    ttk.Label(
        top,
        text="作業フロー管理 (Workflow Manager)",
        font=UFont(14, "bold")
    ).pack(side="left", padx=8)

    # Steps/info card
    step_card = ttk.Frame(outer, style="Card.TFrame")
    step_card.pack(fill="x", pady=(4, 8))
    ttk.Label(
        step_card,
        text="手順",
        font=UFont(11, "bold")
    ).pack(anchor="w", pady=(4, 0), padx=8)
    ttk.Label(
        step_card,
        text="① カレンダーページを選択",
        font=UFont(10)
    ).pack(anchor="w", padx=16, pady=1)
    ttk.Label(
        step_card,
        text="② Excelファイルを選択",
        font=UFont(10)
    ).pack(anchor="w", padx=16, pady=1)
    ttk.Label(
        step_card,
        text="③ 製品を確認して同期",
        font=UFont(10)
    ).pack(anchor="w", padx=16, pady=1)

    hint = ttk.Label(
        outer,
        text="上から順番に ① → ② → ③ のボタンを押すだけでOKです。",
        font=UFont(9),
        foreground=THEMES[CURRENT_THEME]["accent"],
    )
    hint.pack(anchor="w", pady=(0, 4))

    steps = ttk.Frame(outer)
    steps.pack(fill="x", pady=(4, 0))
    ttk.Button(
        steps,
        text="１ カレンダーページを選択",
        style="Menu.TButton",
        command=select_calendar_for_wf
    ).pack(pady=3, fill="x")
    ttk.Button(
        steps,
        text="２ Excel ファイルを選択",
        style="Menu.TButton",
        command=select_excel_for_wf
    ).pack(pady=3, fill="x")
    ttk.Button(
        steps,
        text="３ 製品一覧を表示して同期",
        style="Menu.TButton",
        command=show_workflow_products_screen,
    ).pack(pady=3, fill="x")

    # Status summary
    summary = ttk.Frame(outer)
    summary.pack(pady=(12, 0), fill="x")

    cal_txt = wf.selected_page_title or "未接続"
    excel_txt = wf.selected_file or "未選択"

    ttk.Label(
        summary,
        text=f"📅 カレンダー: {cal_txt}",
        font=UFont(10, "bold")
    ).pack(anchor="w")

    ttk.Label(
        summary,
        text=f"📂 Excel: {excel_txt}",
        font=UFont(10)
    ).pack(anchor="w")

    set_status("作業フローモード：手順 １ → ２ → ３ の順に進んでください。")


def select_calendar_for_wf():
    """Show a simple dialog listing the next few calendar pages and let the user pick one."""
    import logic
    
    pages = logic.get_calendar_pages_next4()
    if not pages:
        messagebox.showinfo("カレンダー", "日付の近いカレンダーページが見つかりませんでした。")
        return

    dialog = tk.Toplevel(root)
    dialog.title("カレンダーページを選択")
    dialog.grab_set()

    ttk.Label(
        dialog,
        text="使用するカレンダーページを選んでください。",
        font=(UI_FONT_FAMILY, 11, "bold")
    ).pack(padx=12, pady=(12, 4))

    var = tk.StringVar()
    for pid, title, d in pages:
        label = f"{d}  :  {title}"
        rb = ttk.Radiobutton(dialog, text=label, value=pid, variable=var)
        rb.pack(anchor="w", padx=16, pady=2)

    def on_ok():
        pid = var.get()
        if not pid:
            messagebox.showwarning("選択", "ページを選んでください。")
            return
        # Save into workflow_manager
        for pid_, title, d in pages:
            if pid_ == pid:
                wf.selected_page_id = pid
                wf.selected_page_title = title
                wf.wf_progress["calendar_selected"] = True
                break
        set_status(f"カレンダー選択: {wf.selected_page_title}")
        ui_log(f"[WF] Selected calendar page: {wf.selected_page_title} ({pid})")
        dialog.destroy()
        show_workflow_manager_home()

    ttk.Button(dialog, text="OK", style="Accent.TButton", command=on_ok).pack(pady=(8, 4))
    ttk.Button(dialog, text="キャンセル", style="Menu.TButton", command=dialog.destroy).pack(pady=(0, 12))


def select_excel_for_wf():
    """Let the user pick an Excel file.

    Always prefer the folder chosen in Main Menu (saved in settings).
    After selecting a file, persist its folder so all modules stay in sync.
    """
    from tkinter import filedialog

    # Prefer user-selected folder (settings.json), then current logic.FOLDER_PATH
    saved = _get_excel_folder_pref()
    initial_dir = _normalize_dialog_dir(saved)
    if not initial_dir:
        initial_dir = os.path.expanduser("~")

    ui_log(f"[WF] Excel picker initialdir = {initial_dir}")

    fname = filedialog.askopenfilename(
        parent=root,
        title="Excel（塗装）ファイルを選択",
        initialdir=initial_dir,
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"), ("All files", "*.*")],
    )
    if not fname:
        return

    # Persist folder so next dialogs open in the same place
    try:
        picked_dir = os.path.dirname(fname)
        if picked_dir:
            _set_excel_folder_pref(picked_dir)
    except Exception:
        pass

    wf.selected_file = fname
    wf.wf_progress["excel_selected"] = True
    set_status(f"Excel選択: {fname}")
    ui_log(f"[WF] Selected Excel: {fname}")
    show_workflow_manager_home()



 
def _wf_run_sync_safe():
    """Run Workflow sync with visible logging + error popup."""
    try:
        ui_log("[WF] Sync button pressed")
        ui_log(f"[WF] selected_page_id = {getattr(wf, 'selected_page_id', None)}")
        ui_log(f"[WF] selected_file    = {getattr(wf, 'selected_file', None)}")
        wf.highlight_and_sync()
        ui_log("[WF] highlight_and_sync() finished")
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        ui_log("[WF][ERROR] highlight_and_sync crashed")
        ui_log(tb)
        try:
            messagebox.showerror("同期エラー", f"同期処理でエラーが発生しました。\n\n{e}\n\n詳細:\n{tb}")
        except Exception:
            pass

def select_excel_for_daily_ui():
    """Let the user pick an Excel file for Daily Workflow Generator.

    Uses the shared Excel folder preference from the Main Menu.
    After selecting a file, persists its folder so dialogs stay in sync.
    Stores the chosen file into daily_workflow_generator.selected_excel_file.
    """
    from tkinter import filedialog

    saved = _get_excel_folder_pref()
    initial_dir = _normalize_dialog_dir(saved)
    if not initial_dir:
        initial_dir = os.path.expanduser("~")

    ui_log(f"[DailyUI] Excel picker initialdir = {initial_dir}")

    fname = filedialog.askopenfilename(
        parent=root,
        title="Daily: Excel ファイルを選択",
        initialdir=initial_dir,
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"), ("All files", "*.*")],
    )
    if not fname:
        return

    # Persist folder so next dialogs open in the same place
    try:
        picked_dir = os.path.dirname(fname)
        if picked_dir:
            _set_excel_folder_pref(picked_dir)
    except Exception:
        pass

    # Save into Daily module for this session
    try:
        daily.selected_excel_file = fname
    except Exception:
        pass

    set_status(f"Daily Excel選択: {fname}")
    ui_log(f"[DailyUI] Selected Excel: {fname}")

    # Re-render Daily screen so the label reflects the selection
    try:
        show_daily_generator_home()
    except Exception:
        # Fallback: call current screen callback if available
        try:
            if _current_screen_cb:
                _current_screen_cb()
        except Exception:
            pass
# =====================================================
# WORKFLOW PRODUCTS SCREEN (big checkbox list)
# =====================================================

def show_workflow_products_screen():
    """
    Detailed Workflow Manager screen that shows the per-part checkbox list,
    similar to the original wfm_v1114 UI.

    Uses wf.load_products_from_excel() to fill wf.checkbox_vars and then
    renders them in a scrollable list grouped by color.
    """
    if not wf.selected_page_id:
        messagebox.showwarning("Warning", "先にカレンダーページを選択してください。")
        return
    if not wf.selected_file:
        messagebox.showwarning("Warning", "先に Excel ファイルを選択してください。")
        return

    clear_main_frame()
    set_current_screen(show_workflow_products_screen)

    outer = ttk.Frame(main_frame)
    outer.pack(fill="both", expand=True, padx=16, pady=16)

    # Top row: back + title
    top = ttk.Frame(outer)
    top.pack(fill="x", pady=(0, 8))

    ttk.Button(
        top,
        text="← 戻る",
        style="Menu.TButton",
        command=show_workflow_manager_home,
    ).pack(side="left")

    ttk.Label(
        top,
        text="製品一覧（チェックして同期）",
        font=UFont(18, "bold"),
    ).pack(side="left", padx=8)

    # Summary line
    summary = ttk.Frame(outer)
    summary.pack(fill="x", pady=(4, 8))

    cal_txt = wf.selected_page_title or "未接続"
    excel_txt = wf.selected_file or "未選択"

    ttk.Label(
        summary,
        text=f"📅 カレンダー: {cal_txt}",
        font=UFont(12, "bold"),
    ).pack(anchor="w")
    ttk.Label(
        summary,
        text=f"📂 Excel: {excel_txt}",
        font=UFont(12),
    ).pack(anchor="w")

    # Body: scrollable list
    body = ttk.Frame(outer)
    body.pack(fill="both", expand=True, pady=(4, 8))

    # Canvas + scrollbar for scrollable product list
    # Use the app theme background to avoid a white strip on the side
    canvas_bg = THEMES[CURRENT_THEME]["bg"]
    canvas = tk.Canvas(body, highlightthickness=0, bd=0, bg=canvas_bg)

    vscroll = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vscroll.set)

    canvas.pack(side="left", fill="both", expand=True)
    vscroll.pack(side="right", fill="y")

    inner = ttk.Frame(canvas)
    # Base padding; columns/padding will be handled inside, but we keep a modest margin
    inner.configure(padding=(16, 0))
    inner_window = canvas.create_window((0, 0), window=inner, anchor="nw")

    # Update scrollregion when contents change
    def _on_inner_config(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    inner.bind("<Configure>", _on_inner_config)

    # Make the inner frame follow the canvas width to avoid a blank strip on the side
    def _on_canvas_config(event):
        canvas.itemconfigure(inner_window, width=event.width)
    canvas.bind("<Configure>", _on_canvas_config)

    # Optionally allow mousewheel scroll
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # Load products into wf.checkbox_vars
    rows = wf.load_products_from_excel()
    if not rows:
        ttk.Label(
            inner,
            text="Excel から製品が読み込めませんでした。",
            font=(UI_FONT_FAMILY, 10),
        ).pack(pady=8)
        set_status("製品がありません。", THEMES[CURRENT_THEME]["status_warn"])
        return

    # Group rows by color for nicer layout
    by_color: dict[str, list[tuple]] = {}
    for tup in rows:
        (mb, ac, ob, trial, part, color, qty, ct, d) = tup
        by_color.setdefault(color, []).append(tup)

    # Prepare inner grid for 2-column layout of color groups (equal-width columns)
    inner.grid_columnconfigure(0, weight=1, uniform="colorcols")
    inner.grid_columnconfigure(1, weight=1, uniform="colorcols")

    groups_sorted = sorted(by_color.items(), key=lambda x: x[0])

    # Build UI per color group (2 columns, card style)
    # Unified icon style for main and override checkboxes
    ICON_ON = "✅"
    ICON_OFF = "⬜️"
    ICON_FONT = UFont(22, "bold")
    for idx, (color, items) in enumerate(groups_sorted):
        color_key = color

        group_frame = ttk.Frame(
            inner,
            style="Card.TFrame",
            borderwidth=1,
            relief="solid",
        )

        col = idx % 2
        row = idx // 2

        group_frame.grid(
            row=row,
            column=col,
            sticky="nwe",
            padx=8,
            pady=(8, 12),
        )
        group_frame.configure(padding=(8, 4))

        # Header row: icon, color name, 全チェック切替 button
        header = ttk.Frame(group_frame)
        header.pack(fill="x", pady=(2, 1), padx=8)

        icon_label = ttk.Label(header, text="⬜")
        icon_label.pack(side="left", padx=(0, 6))

        color_label = ttk.Label(
            header,
            text=f"{color_key}",
            font=UFont(20, "bold"),
            foreground=THEMES[CURRENT_THEME]["accent"],
        )
        color_label.pack(side="left")

        # 全チェック切替 button on the right (toggles all MB for this color)
        ttk.Button(
            header,
            text="全チェック切替",
            style="Menu.TButton",
            command=lambda c=color_key: wf.toggle_sync_for_color_group(c),
        ).pack(side="right")

        # Helper MB/OB line
        ttk.Label(
            group_frame,
            text="MB: 通常同期   /   OB: Excelの状態を無視して上書き",
            font=UFont(9),
            foreground=THEMES[CURRENT_THEME]["accent_fg"],
        ).pack(anchor="w", padx=8, pady=(0, 2))

        # Combined table for checkbox headers + rows
        table = ttk.Frame(group_frame)
        table.pack(fill="x", padx=8, pady=(0, 4))

        # Columns: 0=同期, 1=色付き, 2=上書き, 3=部品名, 4=数量/c-t
        table.grid_columnconfigure(0, weight=0)
        table.grid_columnconfigure(1, weight=0)
        table.grid_columnconfigure(2, weight=0)
        table.grid_columnconfigure(3, weight=1)
        table.grid_columnconfigure(4, weight=0)

        # Header row
        ttk.Label(table, text="同期", font=UFont(9)).grid(
            row=0, column=0, padx=(12, 4), pady=(0, 2), sticky="w"
        )
        ttk.Label(table, text="色付き", font=UFont(9)).grid(
            row=0, column=1, padx=(12, 4), pady=(0, 2), sticky="w"
        )
        ttk.Label(table, text="上書き", font=UFont(9)).grid(
            row=0, column=2, padx=(12, 4), pady=(0, 2), sticky="w"
        )

        row_font = UFont(13)
        qty_font = UFont(10, "bold")

        # Data rows
        for row_index, (mb, ac, ob, trial, part, color_val, qty, ct, d) in enumerate(items, start=1):
            # --- MB checkbox icon ---
            # --- MB checkbox icon ---
            mb_icon = ttk.Label(
                table,
                text=ICON_ON if mb.get() else ICON_OFF,
                width=2,
                anchor="center",
                font=ICON_FONT,
            )
            mb_icon.grid(row=row_index, column=0, padx=(12, 4), pady=(0, 0), sticky="w")

            # register for group toggle (header icon + row icon)
            wf.color_checkbox_icon_map.setdefault(color_key, []).append(
                (mb, icon_label, mb_icon)
            )

            def _toggle_mb(var=mb, icon=mb_icon, color_key=color_key):
                new_val = not var.get()
                var.set(new_val)
                icon.config(text=ICON_ON if new_val else ICON_OFF)

            mb_icon.bind("<Button-1>", lambda e, fn=_toggle_mb: fn())

            # --- AC checkbox icon ---
            ac_icon = ttk.Label(
                table,
                text=ICON_ON if ac.get() else ICON_OFF,
                width=2,
                anchor="center",
                font=ICON_FONT,
            )
            ac_icon.grid(row=row_index, column=1, padx=(12, 4), pady=(0, 0), sticky="w")

            def _toggle_ac(var=ac, icon=ac_icon):
                new_val = not var.get()
                var.set(new_val)
                icon.config(text=ICON_ON if new_val else ICON_OFF)

            ac_icon.bind("<Button-1>", lambda e, fn=_toggle_ac: fn())

            # --- OB checkbox icon ---
            ob_icon = ttk.Label(
                table,
                text=ICON_ON if ob.get() else ICON_OFF,
                width=2,
                anchor="center",
                font=ICON_FONT,
            )
            ob_icon.grid(row=row_index, column=2, padx=(12, 4), pady=(0, 0), sticky="w")


            def _toggle_ob(var=ob, icon=ob_icon):
                new_val = not var.get()
                var.set(new_val)
                icon.config(text=ICON_ON if new_val else ICON_OFF)

            ob_icon.bind("<Button-1>", lambda e, fn=_toggle_ob: fn())

            # --- Part name (trial + part, wrapping) ---
            if trial:
                main_text = f"{trial}   {part}"
            else:
                main_text = part

            part_label = ttk.Label(
                table,
                text=main_text,
                font=row_font,
                anchor="w",
                justify="left",
            )
            part_label.grid(row=row_index, column=3, padx=(4, 8), pady=(0, 0), sticky="w")

            # Adjust wraplength when the row resizes
            def _on_row_configure(event, lbl=part_label):
                try:
                    available = max(80, event.width - 260)
                    lbl.configure(wraplength=available)
                except Exception:
                    pass

            part_label.grid(row=row_index, column=3, padx=(4, 8), pady=(0, 0), sticky="w")

            # --- Right side: 数量 / c-t summary ---
            qty_ct_text = f"数量: {qty}   c/t: {ct}"
            qty_label = ttk.Label(
                table,
                text=qty_ct_text,
                font=qty_font,
                anchor="e",
            )
            qty_label.grid(row=row_index, column=4, padx=(4, 2), pady=(0, 0), sticky="e")

    # Bottom controls: sync button
    bottom = ttk.Frame(outer)
    bottom.pack(fill="x", pady=(8, 0))

    ttk.Button(
        bottom,
        text="選択した製品を同期",
        style="Accent.TButton",
        command=_wf_run_sync_safe,
    ).pack(side="right")

    ui_log("[WF] Products screen ready")
    set_status("製品一覧：チェックを確認して同期してください。")


# =====================================================
# DAILY GENERATOR SCREEN
# =====================================================

_selected_page_for_daily: dict | None = None  # {"id": ..., "title": ..., "date": ...}

def show_daily_generator_home():
    clear_main_frame()
    set_current_screen(show_daily_generator_home)

    outer = ttk.Frame(main_frame)
    outer.pack(fill="both", expand=True, padx=8, pady=16)

    # Top
    top = ttk.Frame(outer)
    top.pack(fill="x", pady=(0, 8))

    ttk.Button(
        top,
        text="← メニューへ",
        style="Menu.TButton",
        command=show_main_menu
    ).pack(side="left")

    ttk.Label(
        top,
        text="Daily Workflow Generator",
        font=(UI_FONT_FAMILY, 14, "bold")
    ).pack(side="left", padx=8)

        # --- Daily: manual Excel selection (optional) ---
    daily_excel_card = ttk.Frame(outer, style="Card.TFrame")
    daily_excel_card.pack(fill="x", pady=(6, 8))

    ttk.Label(
        daily_excel_card,
        text="Excel（任意）",
        font=UFont(11, "bold"),
    ).pack(anchor="w", padx=10, pady=(6, 0))

    cur_daily_excel = getattr(daily, "selected_excel_file", None) or "未選択"
    ttk.Label(
        daily_excel_card,
        text=f"📄 選択中: {cur_daily_excel}",
        font=UFont(9),
        wraplength=720,
        justify="left",
    ).pack(anchor="w", padx=14, pady=(2, 6))

    ttk.Button(
        daily_excel_card,
        text="Excel ファイルを選択",
        style="Menu.TButton",
        command=select_excel_for_daily_ui,
    ).pack(anchor="e", padx=12, pady=(0, 10))

    # Info/explanation card
    info_card = ttk.Frame(outer, style="Card.TFrame")
    info_card.pack(fill="x", pady=(4, 8))
    ttk.Label(
        info_card,
        text="カレンダーページを選択して、その日の Excel から\n自動で『作業内容』を生成します。\n黄色ハイライト済みの行はスキップされます。",
        justify="left",
        font=(UI_FONT_FAMILY, 10),
    ).pack(anchor="w", padx=8, pady=8)

    controls = ttk.Frame(outer)
    controls.pack(pady=4, fill="x")
    ttk.Button(
        controls,
        text="📅 カレンダーページを選択",
        style="Menu.TButton",
        command=select_calendar_for_daily
    ).pack(pady=3, fill="x")
    ttk.Button(
        controls,
        text="▶ Daily Generator 実行",
        style="Menu.TButton",
        command=run_daily_for_selected,
    ).pack(pady=3, fill="x")

    # Summary
    summary = ttk.Frame(outer)
    summary.pack(pady=(12, 0), fill="x")

    title = _selected_page_for_daily["title"] if _selected_page_for_daily else "未選択"
    d = _selected_page_for_daily["date"] if _selected_page_for_daily else "----/--/--"

    ttk.Label(
        summary,
        text=f"📅 カレンダー: {title}",
        font=(UI_FONT_FAMILY, 10, "bold")
    ).pack(anchor="w")
    ttk.Label(
        summary,
        text=f"日付: {d}",
        font=(UI_FONT_FAMILY, 10)
    ).pack(anchor="w")

    set_status("Daily Generator：カレンダーページを選んでください。")


def select_calendar_for_daily():
    global _selected_page_for_daily

    pages = logic.get_calendar_pages_next4()
    if not pages:
        messagebox.showinfo("カレンダー", "日付の近いカレンダーページが見つかりませんでした。")
        return

    dialog = tk.Toplevel(root)
    dialog.title("Daily用 カレンダーページ選択")
    dialog.grab_set()

    ttk.Label(
        dialog,
        text="Daily Generator に使用するカレンダーページを選んでください。",
        font=(UI_FONT_FAMILY, 11, "bold")
    ).pack(padx=12, pady=(12, 4))

    var = tk.StringVar()
    # mapping: key = string page id (for Tk), value = (original_id, title, date)
    mapping = {}
    for pid, title, d in pages:
        pid_str = str(pid)
        label = f"{d}  :  {title}"
        mapping[pid_str] = (pid, title, d)
        ttk.Radiobutton(
            dialog,
            text=label,
            value=pid_str,
            variable=var,
        ).pack(anchor="w", padx=16, pady=2)

    def on_ok():
        global _selected_page_for_daily
        pid_str = var.get()
        if not pid_str:
            messagebox.showwarning("選択", "ページを選んでください。")
            return

        if pid_str not in mapping:
            messagebox.showerror("選択", "内部エラー：ページ情報の取得に失敗しました。もう一度選択してください。")
            return

        original_id, title, d = mapping[pid_str]
        _selected_page_for_daily = {"id": original_id, "title": title, "date": d}
        ui_log(f"[Daily] Selected calendar page: {title} ({d})")
        dialog.destroy()
        show_daily_generator_home()

    ttk.Button(dialog, text="OK", style="Accent.TButton", command=on_ok).pack(pady=(8, 4))
    ttk.Button(dialog, text="キャンセル", style="Menu.TButton", command=dialog.destroy).pack(pady=(0, 12))


def run_daily_for_selected():
    if not _selected_page_for_daily:
        messagebox.showwarning("Daily", "先にカレンダーページを選択してください。")
        return

    pid = _selected_page_for_daily["id"]
    d_str = _selected_page_for_daily["date"]

    try:
        target_date = datetime.fromisoformat(d_str).date()
    except Exception:
        # Fallback: try to parse manually if needed
        try:
            target_date = datetime.strptime(d_str, "%Y-%m-%d").date()
        except Exception:
            messagebox.showerror("Daily", f"日付の解析に失敗しました: {d_str}")
            return

    # daily.run_daily_auto_for_page only needs a dict with "id"
    selected_page_stub = {"id": pid}

    set_status("Daily Generator 実行中…", THEMES[CURRENT_THEME]["accent"])
    set_progress(0, mode="determinate")
    ui_log(f"[Daily] Running for date {target_date} on page {pid}")

    daily.run_daily_auto_for_page(selected_page_stub, target_date)


# =====================================================
# APP INIT
# =====================================================

def build_root():
    global root, main_frame, status_label, progress, log_console

    root = tk.Tk()
    # --- Full Custom Dark Theme ---
    style = ttk.Style()
    style.theme_use("clam")

    # Base colors (Notion-like dark)
    BG_MAIN = "#1B1B1B"      # main background
    BG_PANEL = "#202020"     # cards / panels
    BG_BUTTON = "#262626"    # default button background
    FG_TEXT = "#FFFFFF"      # primary text
    BORDER_COLOR = "#2E2E2E" # subtle borders
    ACCENT = "#4B93FF"       # Notion-like blue

    # Frame / Containers
    style.configure("TFrame", background=BG_MAIN, bordercolor=BORDER_COLOR)
    style.configure("Dark.TFrame", background=BG_MAIN, bordercolor=BORDER_COLOR)
    style.configure("Card.TFrame", background=BG_PANEL, bordercolor=BORDER_COLOR)

    # Labels
    style.configure("TLabel", background=BG_MAIN, foreground=FG_TEXT)
    style.configure("Dark.TLabel", background=BG_MAIN, foreground=FG_TEXT)

    # Default Buttons (sharp edges)
    style.configure(
        "TButton",
        background=BG_BUTTON,
        foreground=FG_TEXT,
        padding=6,
        bordercolor=BORDER_COLOR,
    )
    style.map(
        "TButton",
        background=[("active", "#333333")]
    )

    # Menu buttons (サブ動作用)
    style.configure(
        "Menu.TButton",
        background=BG_PANEL,
        foreground=FG_TEXT,
        padding=6,
        bordercolor=BORDER_COLOR,
    )
    style.map(
        "Menu.TButton",
        background=[("active", "#2A2A2A")]
    )

    # Accent buttons (メイン操作用)
    style.configure(
        "Accent.TButton",
        background=ACCENT,
        foreground=FG_TEXT,
        padding=6,
        bordercolor=ACCENT,
    )
    style.map(
        "Accent.TButton",
        background=[("active", "#3C7AD9")]
    )

    # Entries / Combobox
    style.configure("TEntry", fieldbackground=BG_BUTTON, foreground=FG_TEXT)
    style.configure("TCombobox", fieldbackground=BG_BUTTON, foreground=FG_TEXT)

    # Scrollbars
    style.configure("Vertical.TScrollbar", background=BG_BUTTON, troughcolor=BG_PANEL, bordercolor=BORDER_COLOR)
    style.configure("Horizontal.TScrollbar", background=BG_BUTTON, troughcolor=BG_PANEL, bordercolor=BORDER_COLOR)

    # Notebook / Tabs
    style.configure("TNotebook", background=BG_MAIN, bordercolor=BORDER_COLOR)
    style.configure("TNotebook.Tab", background=BG_BUTTON, foreground=FG_TEXT, padding=[8, 4])
    style.map("TNotebook.Tab", background=[("selected", BG_PANEL)])

    # Treeview
    style.configure(
        "Treeview",
        background=BG_PANEL,
        foreground=FG_TEXT,
        fieldbackground=BG_PANEL,
        bordercolor=BORDER_COLOR,
    )
    style.map(
        "Treeview",
        background=[("selected", "#333333")]
    )

    # Radiobuttons / Checkbuttons
    style.configure(
        "TRadiobutton",
        background=BG_MAIN,
        foreground=FG_TEXT,
    )
    style.configure(
        "TCheckbutton",
        background=BG_MAIN,
        foreground=FG_TEXT,
    )
    # END Dark Theme
    root.title("Painting Team App (Modular)")
    root.geometry("1000x720")
    root.configure(bg="#1B1B1B")

    ensure_dpi_awareness_and_scaling()
    apply_theme(root)

    # Main vertical layout: top bar, content, log, status
    root.rowconfigure(1, weight=1)
    root.columnconfigure(0, weight=1)

    # Top bar
    topbar = ttk.Frame(root)
    topbar.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))

    ttk.Label(
        topbar,
        text="塗装チームアプリ",
        font=(UI_FONT_FAMILY, 13, "bold")
    ).pack(side="left", padx=(0, 8))

    ver_label = ttk.Label(
        topbar,
        text=f"v{APP_VERSION}",
        font=(UI_FONT_FAMILY, 9)
    )
    ver_label.pack(side="right", padx=(0, 8))

    ttk.Button(
        topbar,
        text="テーマ切替",
        style="Accent.TButton",
        command=lambda: toggle_theme(lambda: apply_theme(root))
    ).pack(side="right")

    # Main content frame
    main_frame = ttk.Frame(root)
    main_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)

    # Log console
    log_frame = ttk.Frame(root)
    log_frame.grid(row=2, column=0, sticky="ew", padx=8, pady=(4, 2))

    ttk.Label(log_frame, text="ログ:").pack(anchor="w")
    log_console = tk.Text(log_frame, height=8)
    log_console.pack(fill="both", expand=True)
    log_console.configure(bg="#202020", fg="#FFFFFF", insertbackground="#FFFFFF", bd=0, highlightthickness=0)

    # Status bar
    status_frame = ttk.Frame(root)
    status_frame.grid(row=3, column=0, sticky="ew", padx=8, pady=(2, 8))

    status_label = ttk.Label(
        status_frame,
        text="Ready.",
        anchor="w"
    )
    status_label.pack(side="left", fill="x", expand=True)

    progress = ttk.Progressbar(status_frame, mode="determinate", length=180)
    progress.pack(side="right", padx=(8, 0))

    # Wire shared widgets into modules
    wf.main_frame = main_frame
    wf.status_label = status_label
    wf.log_console = log_console
    wf.progress = progress

    daily.status_label = status_label
    daily.progress = progress

    # Hook backend logger to UI
    logic.set_log_handler(ui_log)

    # Initial screen
    show_main_menu()


def main():
    build_root()
    root.mainloop()


if __name__ == "__main__":
    main()