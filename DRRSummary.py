import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import json
import os

# --- Configuration & Styling ---
CONFIG_FILE = "config.json"
BG_COLOR = "#243447"  # Matches your dark theme
BOX_BG = "#2c3e50"
TEXT_COLOR = "#1abc9c" # The green accent color

class PredictiveMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("CMS Predictive Monitor Dashboard")
        self.root.geometry("1000x800")
        self.root.configure(bg=BG_COLOR)

        # 💾 Persistence: Load the last used file
        self.master_file_path = self.load_config()
        self.search_var = tk.StringVar()
        
        # Dictionary to hold the 9 metric variables for easy updating
        self.metrics = {
            "Calls Dialed": tk.StringVar(value="0"), "Call Ringing": tk.StringVar(value="0"),
            "Call Waiting": tk.StringVar(value="0"), "Logged In": tk.StringVar(value="0"),
            "In Calls": tk.StringVar(value="0"), "Waiting": tk.StringVar(value="0"),
            "Paused": tk.StringVar(value="0"), "In Dead Calls": tk.StringVar(value="0"),
            "In Dispo": tk.StringVar(value="0")
        }

        self.setup_ui()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                return json.load(f).get("last_file")
        return None

    def save_config(self, path):
        with open(CONFIG_FILE, "w") as f:
            json.dump({"last_file": path}, f)

    def setup_ui(self):
        # 📂 Top Bar: File Upload and Search
        header = tk.Frame(self.root, bg=BG_COLOR)
        header.pack(fill="x", pady=10)

        btn_upload = tk.Button(header, text="📂 UPLOAD MASTER XLSX", command=self.upload_file)
        btn_upload.pack()

        self.lbl_file = tk.Label(header, text=f"Master: {os.path.basename(self.master_file_path or 'None')}", 
                                 bg=BG_COLOR, fg=TEXT_COLOR)
        self.lbl_file.pack()

        search_frame = tk.Frame(self.root, bg=BG_COLOR)
        search_frame.pack(pady=5)
        
        tk.Entry(search_frame, textvariable=self.search_var, width=30).pack(side="left", padx=5)
        # 🖱️ Trigger: Enter key starts search
        self.root.bind('<Return>', lambda e: self.refresh_data())
        tk.Button(search_frame, text="🔍 SEARCH", command=self.refresh_data).pack(side="left")

        # 🖼️ Main Content: Metrics Grid and Side Panel
        main_body = tk.Frame(self.root, bg=BG_COLOR)
        main_body.pack(fill="both", expand=True, padx=20)

        # Left: 3x3 Metrics Grid
        grid_frame = tk.Frame(main_body, bg=BG_COLOR)
        grid_frame.pack(side="left", fill="both", expand=True)

        for i, (label, var) in enumerate(self.metrics.items()):
            box = tk.Frame(grid_frame, bg=BG_COLOR, highlightbackground=TEXT_COLOR, highlightthickness=1)
            box.grid(row=i//3, column=i%3, padx=10, pady=10, sticky="nsew")
            
            tk.Label(box, text=label, bg=BG_COLOR, fg="white", font=("Arial", 10)).pack(pady=5)
            tk.Label(box, textvariable=var, bg=BG_COLOR, fg=TEXT_COLOR, font=("Arial", 18, "bold")).pack(pady=10)

        # 🔴 Right: Missing Agents Panel
        side_panel = tk.Frame(main_body, bg="#1a252f", width=200)
        side_panel.pack(side="right", fill="y", padx=10)
        
        tk.Label(side_panel, text="MISSING AGENTS", bg="#1a252f", fg="#e74c3c", font=("Arial", 10, "bold")).pack(pady=10)
        self.missing_list = tk.Listbox(side_panel, bg="#1a252f", fg="#e74c3c", borderwidth=0, highlightthickness=0)
        self.missing_list.pack(fill="both", expand=True)

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.master_file_path = path
            self.save_config(path)
            self.lbl_file.config(text=f"Master: {os.path.basename(path)}")

    def refresh_data(self):
        # 1. Logic: Clean API usernames (Split at space + Upper)
        # 2. Logic: Clean Excel usernames (Strip + Upper)
        # 3. Logic: Compare and update self.metrics and self.missing_list
        # Note: Replace 'mock_data' with your specific API call
        print("Refreshing with keyword:", self.search_var.get())
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = PredictiveMonitor(root)
    root.mainloop()