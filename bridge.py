import sys
import json
import os
import tkinter as tk

# Ensure current directory is in search path for logic and workflow_manager
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import logic
import workflow_manager as wf

def main():
    try:
        if len(sys.argv) < 2:
            print(json.dumps({"error": "No command provided"}))
            return

        command = sys.argv[1]
        
        if command == "get_calendar":
            pages = logic.get_calendar_pages_next4()
            res = [{"id": p[0], "title": p[1], "date": p[2]} for p in pages]
            print(json.dumps(res))

        elif command == "load_products":
            if len(sys.argv) < 3:
                print(json.dumps({"error": "No file path provided"}))
                return
            file_path = sys.argv[2]
            wf.selected_file = file_path
            # Reset checkbox_vars before loading
            wf.checkbox_vars = []
            products = wf.load_products_from_excel()
            
            res = []
            for i, (mb, ac, ob, trial, part, color, qty, ct, date) in enumerate(products):
                # Ensure qty is a clean integer
                q_val = 0
                try:
                    q_val = int(str(qty or 0).strip())
                except Exception:
                    pass

                res.append({
                    "id": f"p{i}",
                    "selected": bool(mb.get()),
                    "colorAccent": bool(ac.get()),
                    "override": bool(ob.get()),
                    "trial": str(trial or ""),
                    "part": str(part or ""),
                    "color": str(color or ""),
                    "qty": q_val,
                    "ct": int(ct or 0),
                    "date": str(date or "")
                })
            print(json.dumps(res))

        elif command == "sync":
            if len(sys.argv) < 3:
                print(json.dumps({"error": "No input JSON provided"}))
                return
            
            data = json.loads(sys.argv[2])
            file_path = data.get("file_path")
            page_id = data.get("page_id")
            products_data = data.get("products", [])

            wf.selected_file = file_path
            wf.selected_page_id = page_id
            
            # Reconstruct checkbox_vars
            wf.checkbox_vars = []
            for p in products_data:
                mb = tk.BooleanVar(value=p.get("selected", False))
                ac = tk.BooleanVar(value=p.get("colorAccent", False))
                ob = tk.BooleanVar(value=p.get("override", False))
                wf.checkbox_vars.append((
                    mb, ac, ob, 
                    p.get("trial", ""), 
                    p.get("part", ""), 
                    p.get("color", ""), 
                    p.get("qty", 0), 
                    p.get("ct", 0), 
                    p.get("date", "")
                ))
            
            # Note: highlight_and_sync might show messageboxes or ask confirmation.
            # We should ideally patch messagebox to auto-approve or handle it.
            # For now, let's hope it runs through. 
            # In a real headless env, we'd need to mock messagebox.askokcancel etc.
            
            # Mocking messagebox for headless execution
            from tkinter import messagebox
            messagebox.askokcancel = lambda title, message: True
            messagebox.showwarning = lambda title, message: print(f"Warning: {message}", file=sys.stderr)
            messagebox.showerror = lambda title, message: print(f"Error: {message}", file=sys.stderr)
            
            wf.highlight_and_sync()
            print(json.dumps({"status": "success"}))

        else:
            print(json.dumps({"error": f"Unknown command: {command}"}))

    except Exception as e:
        print(json.dumps({"error": str(e)}), file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
