import customtkinter as ctk
from CTkTable import *
from PIL import Image
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sys
import os

# Add backend to path
sys.path.append(os.path.join(os.getcwd(), "backend"))

from backend.app.core.database import init_db, SessionLocal
from backend.app.models.mail import Mail
from backend.app.services.outlook_service import enumerate_all_folders, search_mails, extract_mails
from backend.app.services.analyzer_service import analyze_mail_content
from backend.app.workers.scan_worker import process_mail

# Appearance Settings
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue") # We will override specific colors

# Professional Color Palette
COLORS = {
    "bg": "#0f172a",         # Deep Slate Blue
    "sidebar": "#1e293b",    # Slate 800
    "accent": "#6366f1",     # Indigo 500
    "accent_hover": "#4f46e5",
    "card_bg": "#1e293b",
    "text_main": "#f8fafc",
    "text_dim": "#94a3b8",
    "danger": "#ef4444",
    "success": "#10b981",
    "warning": "#f59e0b"
}

class MailAnalyzerPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Mail Analyzer Pro - Professional Edition")
        self.geometry("1300x850")
        self.configure(fg_color=COLORS["bg"])
        
        # Load Icons
        self.load_assets()
        
        # Set Window Icon
        try:
            self.iconbitmap("assets/icon.ico")
        except Exception as icon_error:
            print(f"Warning: Could not load window icon: {icon_error}")
            pass
            
        # Initialize DB
        init_db()
        
        # Layout Config
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0, fg_color=COLORS["sidebar"])
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(6, weight=1)
        
        # Logo Section
        self.logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.logo_frame.grid(row=0, column=0, padx=20, pady=(30, 40))
        
        try:
            logo_img = ctk.CTkImage(Image.open("assets/icon.png"), size=(40, 40))
            self.logo_icon = ctk.CTkLabel(self.logo_frame, image=logo_img, text="")
            self.logo_icon.pack(side="left", padx=(0, 10))
        except Exception as logo_error:
            print(f"Warning: Could not load logo image: {logo_error}")
            pass
            
        self.logo_text = ctk.CTkLabel(self.logo_frame, text="ANALYZER", font=ctk.CTkFont(size=18, weight="bold"))
        self.logo_text.pack(side="left")
        
        # Navigation Buttons
        self.nav_dashboard = self.create_nav_btn("Dashboard", 1, self.show_dashboard)
        self.nav_explorer = self.create_nav_btn("PST Explorer", 2, self.show_browser)
        self.nav_search = self.create_nav_btn("Global Search", 3, self.show_search)
        self.nav_results = self.create_nav_btn("Analysis Results", 4, self.show_results)
        
        # --- Main Content ---
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=1, padx=30, pady=30, sticky="nsew")
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(0, weight=1)
        
        # --- Footer / Status ---
        self.status_bar = ctk.CTkFrame(self, height=30, fg_color=COLORS["sidebar"], corner_radius=0)
        self.status_bar.grid(row=1, column=0, columnspan=2, sticky="we")
        
        self.status_text = ctk.CTkLabel(self.status_bar, text="System Ready", font=ctk.CTkFont(size=11), text_color=COLORS["text_dim"])
        self.status_text.pack(side="left", padx=20)
        
        self.show_dashboard()

    def load_assets(self):
        # We could load more specific icons here if they existed
        pass

    def create_nav_btn(self, text, row, command):
        btn = ctk.CTkButton(self.sidebar, text=text, height=45, fg_color="transparent", 
                           text_color=COLORS["text_dim"], hover_color="#2d3748", anchor="w",
                           font=ctk.CTkFont(size=14, weight="normal"), command=command)
        btn.grid(row=row, column=0, padx=20, pady=5, sticky="ew")
        return btn

    def set_active_nav(self, active_btn):
        for btn in [self.nav_dashboard, self.nav_explorer, self.nav_search, self.nav_results]:
            btn.configure(fg_color="transparent", text_color=COLORS["text_dim"])
        active_btn.configure(fg_color=COLORS["accent"], text_color=COLORS["text_main"])

    def clear_main(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()

    def show_dashboard(self):
        self.clear_main()
        self.set_active_nav(self.nav_dashboard)
        
        db = SessionLocal()
        total = db.query(Mail).count()
        high_risk = db.query(Mail).filter(Mail.risk_score >= 50).count()
        db.close()
        
        # Header
        header = ctk.CTkLabel(self.main_container, text="Security Dashboard", font=ctk.CTkFont(size=26, weight="bold"), anchor="w")
        header.pack(fill="x", pady=(0, 30))
        
        # Cards Container
        cards_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        cards_frame.pack(fill="x")
        cards_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        self.create_pro_card(cards_frame, "Analyzed Records", str(total), 0, COLORS["accent"])
        self.create_pro_card(cards_frame, "Threats Detected", str(high_risk), 1, COLORS["danger"])
        self.create_pro_card(cards_frame, "Stores Connected", "Calculating...", 2, COLORS["success"])
        
        # Recent Activity Table
        act_lbl = ctk.CTkLabel(self.main_container, text="Recent Security Findings", font=ctk.CTkFont(size=18, weight="bold"), anchor="w")
        act_lbl.pack(fill="x", pady=(40, 15))
        
        self.show_compact_results()

    def create_pro_card(self, parent, title, value, col, accent_color):
        card = ctk.CTkFrame(parent, fg_color=COLORS["card_bg"], corner_radius=15, height=160)
        card.grid(row=0, column=col, padx=10, sticky="nsew")
        card.grid_propagate(False)
        
        # Accent Border/Indicator
        line = ctk.CTkFrame(card, width=4, fg_color=accent_color, corner_radius=2)
        line.pack(side="left", fill="y", padx=(15, 0), pady=15)
        
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        
        t_lbl = ctk.CTkLabel(content, text=title.upper(), font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_dim"])
        t_lbl.pack(anchor="w")
        
        v_lbl = ctk.CTkLabel(content, text=value, font=ctk.CTkFont(size=32, weight="bold"), text_color=COLORS["text_main"])
        v_lbl.pack(anchor="w", pady=(5, 0))

    def show_browser(self):
        self.clear_main()
        self.set_active_nav(self.nav_explorer)
        
        header = ctk.CTkLabel(self.main_container, text="PST & Mailbox Explorer", font=ctk.CTkFont(size=26, weight="bold"), anchor="w")
        header.pack(fill="x", pady=(0, 20))
        
        container = ctk.CTkFrame(self.main_container, fg_color=COLORS["card_bg"], corner_radius=15)
        container.pack(fill="both", expand=True)
        
        # Tree Styling
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#1e293b", foreground="#f8fafc", fieldbackground="#1e293b", borderwidth=0, rowheight=30)
        style.map("Treeview", background=[('selected', '#6366f1')])
        
        self.tree = ttk.Treeview(container, columns=("Items", "Store"), show="tree headings")
        self.tree.heading("#0", text="Folder")
        self.tree.heading("Items", text="Mails")
        self.tree.heading("Store", text="Source")
        self.tree.pack(fill="both", expand=True, padx=20, pady=20)
        
        btn_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(20, 0))
        
        scan_btn = ctk.CTkButton(btn_frame, text="SCAN SELECTED FOLDER", height=45, fg_color=COLORS["accent"], 
                                hover_color=COLORS["accent_hover"], font=ctk.CTkFont(weight="bold"),
                                command=self.start_folder_scan)
        scan_btn.pack(side="right")
        
        threading.Thread(target=self.load_folders_data, daemon=True).start()

    def load_folders_data(self):
        self.status_text.configure(text="Tuning into Outlook...")
        folders = enumerate_all_folders()
        stores = {}
        for f in folders:
            store = f["store"]
            if store not in stores:
                stores[store] = self.tree.insert("", "end", text=store, open=True)
            self.tree.insert(stores[store], "end", text=f["name"], values=(f["item_count"], f["store"]), tags=(f["entry_id"],))
        self.status_text.configure(text=f"Connected to {len(folders)} folders.")

    def start_folder_scan(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Selection", "Please select a folder to scan.")
            return
            
        tags = self.tree.item(selected[0])["tags"]
        if not tags: return
        entry_id = tags[0]
        
        self.status_text.configure(text="Scanning folder... Please wait.")
        threading.Thread(target=self.run_scan, args=(entry_id,), daemon=True).start()

    def run_scan(self, entry_id):
        from backend.app.services.outlook_service import get_folder_by_id
        folder_obj = get_folder_by_id(entry_id)
        if not folder_obj:
            self.after(0, lambda: messagebox.showerror("Error", "Could not access folder."))
            return
            
        mails = extract_mails(folder_obj)
        db = SessionLocal()
        try:
            count = 0
            total_mails = len(mails)
            for mail in mails:
                analysis = analyze_mail_content(mail["body"])
                process_mail(db, {**mail, **analysis})
                count += 1
                self.after(0, lambda c=count, t=total_mails: self.status_text.configure(text=f"Processed {c}/{t} mails..."))
        finally:
            db.close()
        self.after(0, lambda: messagebox.showinfo("Scan Complete", f"Successfully analyzed {len(mails)} mails."))
        self.after(0, self.show_results)

    def show_search(self):
        self.clear_main()
        self.set_active_nav(self.nav_search)
        
        header = ctk.CTkLabel(self.main_container, text="Advanced Forensic Search", font=ctk.CTkFont(size=26, weight="bold"), anchor="w")
        header.pack(fill="x", pady=(0, 30))
        
        search_box = ctk.CTkFrame(self.main_container, fg_color=COLORS["card_bg"], corner_radius=15)
        search_box.pack(fill="x")
        
        self.sub_entry = self.create_input(search_box, "Subject Pattern:", 0)
        self.body_entry = self.create_input(search_box, "Content Keywords:", 1)
        
        s_btn = ctk.CTkButton(search_box, text="EXECUTE SEARCH", height=45, fg_color=COLORS["accent"], font=ctk.CTkFont(weight="bold"), command=self.run_global_search)
        s_btn.pack(pady=(20, 0))

    def create_input(self, parent, label, row):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=40, pady=10)
        ctk.CTkLabel(f, text=label, font=ctk.CTkFont(size=13), text_color=COLORS["text_dim"]).pack(side="left")
        e = ctk.CTkEntry(f, width=400, height=35, fg_color="#0f172a", border_color="#334155")
        e.pack(side="right")
        return e

    def show_results(self, compact=False):
        if not compact:
            self.clear_main()
            self.set_active_nav(self.nav_results)
            header = ctk.CTkLabel(self.main_container, text="Analysis & Detection Log", font=ctk.CTkFont(size=26, weight="bold"), anchor="w")
            header.pack(fill="x", pady=(0, 20))
            
        container = ctk.CTkFrame(self.main_container, fg_color=COLORS["card_bg"], corner_radius=15)
        container.pack(fill="both", expand=True)
        
        db = SessionLocal()
        mails = db.query(Mail).order_by(Mail.risk_score.desc()).limit(50).all()
        db.close()
        
        table_values = [["SCORE", "SUBJECT", "SENDER", "PST SOURCE"]]
        for m in mails:
            table_values.append([f"{m.risk_score}%", m.subject[:40], m.sender, m.pst_source])
            
        table = CTkTable(container, values=table_values, header_color=COLORS["sidebar"], 
                         text_color=COLORS["text_main"], hover_color="#334155")
        table.pack(fill="both", expand=True, padx=15, pady=15)
        
        if not compact:
            ctk.CTkButton(self.main_container, text="EXPORT FORENSIC REPORT (XLSX)", fg_color=COLORS["success"], 
                          font=ctk.CTkFont(weight="bold"), command=self.export_report).pack(pady=(20, 0), side="right")

    def show_compact_results(self):
        self.show_results(compact=True)

    def export_report(self):
        from backend.app.services.report_service import generate_excel_report
        db = SessionLocal()
        path = generate_excel_report(db)
        db.close()
        if path:
            messagebox.showinfo("Export Success", f"Report saved to:\n{path}")

    def run_global_search(self):
        subject = self.sub_entry.get()
        body = self.body_entry.get()
        
        self.status_text.configure(text="Executing global search across all PSTs...")
        
        def worker():
            params = {"subject": subject, "body_contains": body}
            found = search_mails(params)
            
            db = SessionLocal()
            try:
                for mail in found:
                    analysis = analyze_mail_content(mail["body"])
                    process_mail(db, {**mail, **analysis})
            finally:
                db.close()
                
            self.after(0, lambda: messagebox.showinfo("Search Complete", f"Found and analyzed {len(found)} matching mails."))
            self.after(0, self.show_results)
            
        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = MailAnalyzerPro()
    app.mainloop()
