import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import shutil
import subprocess
import sys
import threading
import queue
import re
import argparse
import webbrowser
import main 
import config # Imports absolute paths

SETTINGS_FILE = config.SETTINGS_FILE
CREDENTIALS_FILE = config.CLIENT_SECRET_FILE
TOKEN_FILE = config.TOKEN_FILE

class EmailBotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automated Email Parser")

        try:
            self.root.state('zoomed')
        except tk.TclError:
            self.root.attributes('-zoomed', True)
            
        self.root.minsize(900, 750)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.configure(bg="#f4f5f7") 
        
        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")
            
        self.style.configure("Treeview", font=("Segoe UI", 10), rowheight=28, background="#ffffff", fieldbackground="#ffffff")
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e1e4e8", foreground="#333333")
        self.style.map("Treeview", background=[('selected', '#0078D7')], foreground=[('selected', 'white')])
        
        self.style.configure("TLabelframe", background="#f4f5f7")
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), background="#f4f5f7", foreground="#2c3e50")
        self.style.configure("TFrame", background="#f4f5f7")
        self.style.configure("TLabel", background="#f4f5f7", font=("Segoe UI", 9))
        self.style.configure("TNotebook", background="#f4f5f7")
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[15, 5])
            
        self.settings = self.load_settings()
        self.process = None
        self.log_queue = queue.Queue()
        
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(fill="both", expand=True)

        # --- HEADER SECTION ---
        lbl_title = tk.Label(main_frame, text="Automated Email Parser", font=("Segoe UI", 18, "bold"), bg="#f4f5f7", fg="#1e293b")
        lbl_title.pack(side="top", pady=(0, 2))
        lbl_creator = tk.Label(main_frame, text="Created by: Jason T. Daohog", font=("Segoe UI", 10, "italic"), bg="#f4f5f7", fg="#64748b")
        lbl_creator.pack(side="top", pady=(0, 10))

        repo_url = "https://github.com/Nenosaj/automated_gmail_parser"
        btn_help = tk.Button(main_frame, text="❓ Help & Documentation", command=lambda: webbrowser.open_new(repo_url), bg="#cbd5e1", fg="#0f172a", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=2, cursor="hand2")
        btn_help.pack(side="top", pady=(0, 15))

        # --- GLOBAL ACTION BUTTONS ---
        frame_actions = ttk.Frame(main_frame)
        frame_actions.pack(side="bottom", fill="x", pady=(5, 0))

        tk.Button(frame_actions, text="Save Configuration", command=self.save_settings, bg="#10b981", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=6).pack(side="left")
        
        # Status Indicator
        self.status_var = tk.StringVar(value=" Bot Status: 🔴 Stopped ")
        self.lbl_status = tk.Label(frame_actions, textvariable=self.status_var, font=("Segoe UI", 11, "bold"), fg="#ef4444", bg="#f4f5f7")
        self.lbl_status.pack(side="left", padx=(20, 0))

        self.btn_stop = tk.Button(frame_actions, text="⏹ Stop Continuous Loop", command=self.stop_pipeline, bg="#ef4444", fg="white", disabledforeground="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=6, state="disabled")
        self.btn_stop.pack(side="right", padx=(10, 0))
        self.btn_start = tk.Button(frame_actions, text="▶ Start Continuous Loop", command=self.start_pipeline, bg="#f59e0b", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=6)
        self.btn_start.pack(side="right")

        # --- TABBED INTERFACE ---
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(side="top", fill="both", expand=True, pady=(0, 15))

        self.tab_config = ttk.Frame(self.notebook, padding="10 10 10 10")
        self.tab_logs = ttk.Frame(self.notebook, padding="10 10 10 10")

        self.notebook.add(self.tab_config, text="⚙️ Configuration & Rules")
        self.notebook.add(self.tab_logs, text="📋 System Logs")

        # ==========================================
        # TAB 1: CONFIGURATION & RULES
        # ==========================================
        top_container = ttk.Frame(self.tab_config)
        top_container.pack(side="top", fill="x", pady=(0, 15))

        frame_api = ttk.LabelFrame(top_container, text=" Step 1: Google API Authentication ", padding="10 10 10 10")
        frame_api.pack(side="left", fill="both", expand=True, padx=(0, 10))
        ttk.Label(frame_api, text="Upload Google Cloud credentials:").pack(side="left")
        self.cred_status_var = tk.StringVar()
        self.lbl_cred_status = tk.Label(frame_api, textvariable=self.cred_status_var, font=("Segoe UI", 9, "bold"), bg="#f4f5f7")
        self.lbl_cred_status.pack(side="left", padx=15)
        tk.Button(frame_api, text="Upload JSON", command=self.upload_credentials, bg="#e2e8f0", relief="groove", padx=10).pack(side="right")
        self.check_credentials_status()

        frame_settings = ttk.LabelFrame(top_container, text=" Step 2: Sender Setup ", padding="10 10 10 10")
        frame_settings.pack(side="right", fill="both", expand=True)
        ttk.Label(frame_settings, text="Sender Email:").pack(side="left")
        self.sender_var = tk.StringVar(value=self.settings.get("sender_email", ""))
        ttk.Entry(frame_settings, textvariable=self.sender_var, width=40).pack(side="left", padx=10, fill="x", expand=True)

        frame_rules = ttk.LabelFrame(self.tab_config, text=" Step 3: Daily Email Rules & Schedule ", padding="15 15 15 15")
        frame_rules.pack(side="top", fill="both", expand=True)

        frame_input = ttk.Frame(frame_rules)
        frame_input.pack(fill="x", pady=(0, 10))
        
        ttk.Label(frame_input, text="Subject:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.new_sub_var = tk.StringVar()
        ttk.Entry(frame_input, textvariable=self.new_sub_var, width=30).grid(row=0, column=1, sticky="w", padx=(0, 15))

        ttk.Label(frame_input, text="Recipients:").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.new_rec_var = tk.StringVar()
        ttk.Entry(frame_input, textvariable=self.new_rec_var, width=35).grid(row=0, column=3, sticky="w", padx=(0, 15))
        tk.Label(frame_input, text="* Separate multiple emails with commas", font=("Segoe UI", 8, "italic"), fg="#64748b", bg="#f4f5f7").grid(row=1, column=3, sticky="w", pady=(0, 5))

        ttk.Label(frame_input, text="Time(24h):").grid(row=0, column=4, sticky="w", padx=(0, 5))
        self.new_time_var = tk.StringVar(value="08:00")
        ttk.Entry(frame_input, textvariable=self.new_time_var, width=8).grid(row=0, column=5, sticky="w", padx=(0, 15))

        frame_form_btns = ttk.Frame(frame_rules)
        frame_form_btns.pack(fill="x", pady=(0, 15))
        
        tk.Button(frame_form_btns, text="+ Add New Rule", command=self.add_rule, bg="#10b981", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=15, pady=4).pack(side="left", padx=(0, 10))
        tk.Button(frame_form_btns, text="💾 Update Selected", command=self.update_rule, bg="#3b82f6", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=15, pady=4).pack(side="left", padx=(0, 10))
        tk.Button(frame_form_btns, text="✖ Clear Form", command=self.clear_form, bg="#e2e8f0", fg="#333333", font=("Segoe UI", 9), relief="flat", padx=15, pady=4).pack(side="left")

        columns = ("subject", "recipients", "time")
        
        frame_rule_btns = ttk.Frame(frame_rules)
        frame_rule_btns.pack(side="bottom", fill="x", pady=(5, 0))
        tk.Button(frame_rule_btns, text="🗑️ Remove Selected", command=self.remove_rule, bg="#ef4444", fg="white", font=("Segoe UI", 9), relief="flat", padx=10, pady=2).pack(side="left")
        tk.Button(frame_rule_btns, text="📅 Sync Tasks to Windows Scheduler", command=self.sync_scheduler, bg="#2563eb", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=5).pack(side="right")

        tree_frame = ttk.Frame(frame_rules)
        tree_frame.pack(side="top", fill="both", expand=True, pady=(0, 5)) 
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("subject", text="Email Subject (Trigger)")
        self.tree.heading("recipients", text="Forward To")
        self.tree.heading("time", text="Daily Schedule")
        self.tree.column("subject", width=500, minwidth=300, anchor="w", stretch=True) 
        self.tree.column("recipients", width=250, minwidth=150, anchor="w", stretch=True)
        self.tree.column("time", width=120, minwidth=120, anchor="center", stretch=False)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # ==========================================
        # TAB 2: SYSTEM LOGS
        # ==========================================
        ttk.Label(self.tab_logs, text="Live Output & Error Reporting:", font=("Segoe UI", 10, "bold")).pack(side="top", anchor="w", pady=(0, 5))
        
        self.log_text = tk.Text(self.tab_logs, state="disabled", bg="#0f172a", fg="#10b981", font=("Consolas", 10), relief="flat", wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll = ttk.Scrollbar(self.tab_logs, command=self.log_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.populate_tree()
        self.update_logs()

    # ------------------------------------------------------------------------
    # LOGIC METHODS 
    # ------------------------------------------------------------------------

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                pass
        return {"sender_email": "", "email_rules": [], "input_dir": config.INPUT_DIR}

    def check_credentials_status(self):
        if os.path.exists(CREDENTIALS_FILE):
            self.cred_status_var.set("✅ Ready")
            self.lbl_cred_status.configure(fg="#10b981")
        else:
            self.cred_status_var.set("❌ Missing")
            self.lbl_cred_status.configure(fg="#ef4444")

    def upload_credentials(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filepath:
            shutil.copy(filepath, CREDENTIALS_FILE)
            self.check_credentials_status()

    def populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for rule in self.settings.get("email_rules", []):
            recs = ", ".join(rule["recipients"])
            time_val = rule.get("time", "08:00")
            self.tree.insert("", "end", values=(rule["subject"], recs, time_val))

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if selected:
            values = self.tree.item(selected[0], "values")
            self.new_sub_var.set(values[0])
            self.new_rec_var.set(values[1])
            self.new_time_var.set(values[2])

    def clear_form(self):
        self.new_sub_var.set("")
        self.new_rec_var.set("")
        self.new_time_var.set("")
        for item in self.tree.selection():
            self.tree.selection_remove(item)

    def validate_inputs(self):
        sub = self.new_sub_var.get().strip()
        recs = self.new_rec_var.get().strip()
        time_val = self.new_time_var.get().strip()
        
        if not sub or not recs:
            messagebox.showwarning("Input Error", "Both Subject and Recipients are required.")
            return None
        if not re.match(r"^([01]\d|2[0-3]):([0-5]\d)$", time_val):
            messagebox.showwarning("Input Error", "Time must be in 24-hour format (e.g., 08:00 or 14:30).")
            return None
        return (sub, recs, time_val)

    def add_rule(self):
        valid_data = self.validate_inputs()
        if valid_data:
            self.tree.insert("", "end", values=valid_data)
            self.clear_form()

    def update_rule(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Select a Rule", "Please click a rule from the table first to update it.")
            return
        valid_data = self.validate_inputs()
        if valid_data:
            self.tree.item(selected[0], values=valid_data)
            self.clear_form()
            self.append_log("[INFO] Rule updated in table. Remember to click 'Save Configuration'.\n")

    def remove_rule(self):
        for item in self.tree.selection():
            self.tree.delete(item)
        self.clear_form()

    def save_settings(self):
        new_sender = self.sender_var.get().strip()
        old_sender = self.settings.get("sender_email", "")

        if new_sender and new_sender != old_sender:
            if os.path.exists(TOKEN_FILE):
                try:
                    os.remove(TOKEN_FILE)
                    self.append_log("[SYSTEM] Sender changed. Cleared old token.json.\n")
                    messagebox.showinfo(
                        "Account Switched", 
                        f"Sender email updated to: {new_sender}\n\n"
                        "Because you changed the Google account, the old login token has been deleted automatically.\n\n"
                        "The next time you start the continuous loop, your browser will open so you can log in and authorize this new account."
                    )
                except Exception as e:
                    self.append_log(f"[ERROR] Could not delete token.json: {e}\n")

        self.settings["sender_email"] = new_sender
        
        new_rules = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            # Convert semicolons to commas for bulletproof inputs
            clean_string = values[1].replace(";", ",")
            recipients_list = [email.strip() for email in clean_string.split(",") if email.strip()]
            new_rules.append({"subject": values[0], "recipients": recipients_list, "time": values[2]})
        
        self.settings["email_rules"] = new_rules
        
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(self.settings, f, indent=4)
            
        if new_sender == old_sender:
            messagebox.showinfo("Success", "Settings saved successfully!")

    def sync_scheduler(self):
        if not os.path.exists(CREDENTIALS_FILE):
            messagebox.showerror("Error", "Please upload your credentials.json first.")
            return
        if not os.path.exists(TOKEN_FILE):
            messagebox.showwarning("Authentication Required", "You must click 'Start Continuous Loop' at least once to log into Google before syncing tasks to the background.")
            return

        self.save_settings() 
        rules = self.settings.get("email_rules", [])
        
        self.append_log("[SCHEDULER] Cleaning up old tasks...\n")
        try:
            find_cmd = ["schtasks", "/query", "/fo", "LIST"]
            result = subprocess.run(find_cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
            existing_tasks = re.findall(r"TaskName:\s+\\(EmailBot_Rule_\d+)", result.stdout)
            
            for task in existing_tasks:
                del_cmd = ["schtasks", "/delete", "/tn", task, "/f"]
                subprocess.run(del_cmd, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                self.append_log(f"[CLEANUP] Removed old task: {task}\n")
        except Exception as e:
            self.append_log(f"[WARNING] Cleanup encountered an issue: {e}\n")

        if not rules:
            messagebox.showinfo("Scheduler Sync", "All tasks have been removed from the Windows Scheduler.")
            return

        success_count = 0
        for i, rule in enumerate(rules):
            subject = rule['subject']
            time_val = rule['time']
            task_name = f"EmailBot_Rule_{i+1}"
            
            if getattr(sys, 'frozen', False):
                action = f'"{sys.executable}" --subject "{subject}"'
            else:
                script_path = os.path.abspath(__file__)
                action = f'"{sys.executable}" "{script_path}" --subject "{subject}"'
            
            try:
                cmd = ["schtasks", "/create", "/tn", task_name, "/tr", action, "/sc", "daily", "/st", time_val, "/f"]
                result = subprocess.run(cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
                
                if result.returncode == 0:
                    success_count += 1
                    self.append_log(f"[SCHEDULER] Successfully scheduled '{subject}' for {time_val} daily.\n")
                else:
                    self.append_log(f"[ERROR] Failed to schedule '{subject}': {result.stderr}\n")
            except Exception as e:
                self.append_log(f"[ERROR] Scheduler error: {e}\n")
                
        messagebox.showinfo("Scheduler Sync", f"Sync complete! {success_count} rule(s) are now active in Windows.")


    def append_log(self, text):
        self.log_queue.put(text)

    def update_logs(self):
        while not self.log_queue.empty():
            line = self.log_queue.get()
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, line)
            self.log_text.see(tk.END)
            self.log_text.config(state="disabled")
        self.root.after(100, self.update_logs)

    def read_process_output(self):
        for line in iter(self.process.stdout.readline, ''):
            self.append_log(line)
        self.process.stdout.close()
        self.process.wait()
        self.append_log("\n[SYSTEM] Continuous Loop stopped.\n")
        
        self.root.after(0, lambda: self.btn_start.config(state="normal", bg="#f59e0b"))
        self.root.after(0, lambda: self.btn_stop.config(state="disabled", bg="#ef4444"))
        self.root.after(0, lambda: self.status_var.set(" Bot Status: 🔴 Stopped "))
        self.root.after(0, lambda: self.lbl_status.config(fg="#ef4444"))
        self.process = None

    def start_pipeline(self):
        if not os.path.exists(CREDENTIALS_FILE):
            messagebox.showerror("Error", "Please upload credentials.json first.")
            return
        if self.process and self.process.poll() is None: return
        self.save_settings()
        
        self.notebook.select(self.tab_logs)
        
        self.append_log("[SYSTEM] Launching continuous loop...\n")
        self.btn_start.config(state="disabled", bg="#e2e8f0")
        self.btn_stop.config(state="normal", bg="#ef4444")
        self.status_var.set(" Bot Status: 🟢 Running ")
        self.lbl_status.config(fg="#10b981")

        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            env = os.environ.copy()
            env["PYTHONUNBUFFERED"] = "1" 
            
            if getattr(sys, 'frozen', False):
                command = [sys.executable, "--loop"]
            else:
                command = [sys.executable, os.path.abspath(__file__), "--loop"]
            
            self.process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, startupinfo=startupinfo, env=env
            )
            threading.Thread(target=self.read_process_output, daemon=True).start()
        except Exception as e:
            self.append_log(f"[ERROR] Failed to start process: {e}\n")

    def stop_pipeline(self):
        if self.process and self.process.poll() is None:
            self.append_log("[SYSTEM] Terminating continuous loop...\n")
            self.process.terminate()

    def on_closing(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
        self.root.destroy()

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--subject", type=str, help="Run a specific rule (For Task Scheduler)")
    parser.add_argument("--loop", action="store_true", help="Run the continuous loop")
    args, unknown = parser.parse_known_args()

    # --- CRITICAL FIX: PREVENT SILENT CRASHES IN BACKGROUND TASKS ---
    if sys.stdout is None or sys.stderr is None:
        log_path = os.path.join(config.BASE_DIR, "background_tasks.log")
        sys.stdout = open(log_path, 'a')
        sys.stderr = sys.stdout

    if args.subject:
        settings = main.load_settings()
        sender = settings.get("sender_email")
        for rule in settings.get("email_rules", []):
            if rule['subject'] == args.subject:
                main.process_rule(rule, sender)
                break
        sys.exit(0)
        
    elif args.loop:
        main.continuous_loop()
        sys.exit(0)
        
    else:
        root = tk.Tk()
        app = EmailBotGUI(root)
        root.mainloop()