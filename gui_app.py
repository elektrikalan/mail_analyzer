import customtkinter as ctk
from PIL import Image
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import sys, os, logging

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from backend.app.core.database import init_db, SessionLocal
from backend.app.models.mail import Mail
from backend.app.services.outlook_service import (
    is_classic_outlook_available, enumerate_stores_tree, extract_mails, search_mails, add_pst_store
)
from backend.app.services.pst_service import (
    is_pypff_available, read_pst_tree, extract_mails_from_pst, search_pst
)
from backend.app.services.analyzer_service import analyze_mail_content
from backend.app.workers.scan_worker import process_mail

logging.basicConfig(level=logging.WARNING)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

BG      = "#0f172a"
SIDE    = "#1e293b"
CARD    = "#1e293b"
ACCENT  = "#6366f1"
AHOVER  = "#4f46e5"
TMAIN   = "#f8fafc"
TDIM    = "#94a3b8"
DANGER  = "#ef4444"
SUCCESS = "#10b981"
WARN    = "#f59e0b"
ROW_ODD = "#162032"
ROW_EVN = "#1e293b"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Mail Analyzer Pro")
        self.geometry("1400x900")
        self.configure(fg_color=BG)
        try:
            self.iconbitmap("assets/icon.ico")
        except Exception:
            pass

        init_db()
        self._pst_files = []          # list of PST paths added by user (PST mode)
        self._mode = self._detect_mode()

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()
        self._build_statusbar()
        self.show_explorer()

    def _detect_mode(self) -> str:
        """Returns 'com' for classic Outlook, 'pst' for new Outlook / no Outlook."""
        if is_classic_outlook_available():
            return "com"
        return "pst"

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color=SIDE)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_rowconfigure(10, weight=1)

        try:
            img = ctk.CTkImage(Image.open("assets/icon.png"), size=(36, 36))
            ctk.CTkLabel(sb, image=img, text="").grid(row=0, column=0, pady=(24, 4))
        except Exception:
            pass

        ctk.CTkLabel(sb, text="MAIL ANALYZER", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=ACCENT).grid(row=1, column=0, padx=16, pady=(0, 20))

        self._nav_btns = {}
        nav_items = [
            ("📁  Explorer",    "explorer",   self.show_explorer),
            ("🔍  Arama",       "search",     self.show_search),
            ("📊  İstatistik",  "stats",      self.show_stats),
            ("📋  Sonuçlar",    "results",    self.show_results),
        ]
        for i, (label, key, cmd) in enumerate(nav_items, start=2):
            btn = ctk.CTkButton(sb, text=label, height=42, fg_color="transparent",
                                text_color=TDIM, hover_color="#2d3748", anchor="w",
                                font=ctk.CTkFont(size=13), command=cmd)
            btn.grid(row=i, column=0, padx=12, pady=3, sticky="ew")
            self._nav_btns[key] = btn
        self._sidebar = sb

    def _set_nav(self, key):
        for k, b in self._nav_btns.items():
            if k == key:
                b.configure(fg_color=ACCENT, text_color=TMAIN)
            else:
                b.configure(fg_color="transparent", text_color=TDIM)

    def _build_main(self):
        self._main = ctk.CTkFrame(self, fg_color="transparent")
        self._main.grid(row=0, column=1, padx=24, pady=24, sticky="nsew")
        self._main.grid_columnconfigure(0, weight=1)
        self._main.grid_rowconfigure(0, weight=1)

    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, height=28, fg_color=SIDE, corner_radius=0)
        bar.grid(row=1, column=0, columnspan=2, sticky="ew")
        self._status = ctk.CTkLabel(bar, text="Hazır", font=ctk.CTkFont(size=11),
                                    text_color=TDIM)
        self._status.pack(side="left", padx=16)

    def _set_status(self, txt):
        self.after(0, lambda: self._status.configure(text=txt))

    def _clear_main(self):
        for w in self._main.winfo_children():
            w.destroy()

    # ──────────────────────────── EXPLORER ────────────────────────────
    def show_explorer(self):
        self._clear_main()
        self._set_nav("explorer")

        # Mode badge
        mode_color = SUCCESS if self._mode == "com" else WARN
        mode_label = "🟢 Classic Outlook (COM)" if self._mode == "com" else "🟡 PST Dosya Modu (Yeni Outlook)"
        hdr = ctk.CTkFrame(self._main, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(hdr, text="PST & Mailbox Explorer",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(side="left")
        ctk.CTkLabel(hdr, text=mode_label, font=ctk.CTkFont(size=11),
                     text_color=mode_color).pack(side="right", padx=4)

        # PST mode: show file adder
        if self._mode == "pst":
            pst_bar = ctk.CTkFrame(self._main, fg_color=CARD, corner_radius=10)
            pst_bar.pack(fill="x", pady=(0, 8))
            ctk.CTkLabel(pst_bar, text="PST dosyası ekle:",
                         text_color=TDIM, font=ctk.CTkFont(size=12)).pack(side="left", padx=12, pady=8)
            ctk.CTkButton(pst_bar, text="+ PST Ekle", width=110, height=32,
                          fg_color=ACCENT, hover_color=AHOVER,
                          command=self._add_pst_file).pack(side="left", padx=4, pady=8)
            self._pst_label = ctk.CTkLabel(pst_bar, text=self._pst_summary(),
                                           text_color=TDIM, font=ctk.CTkFont(size=11))
            self._pst_label.pack(side="left", padx=12)

        pane = tk.PanedWindow(self._main, orient=tk.HORIZONTAL, bg=BG, sashwidth=6, sashrelief="flat")
        pane.pack(fill="both", expand=True)

        # Left: folder tree
        left = ctk.CTkFrame(pane, fg_color=CARD, corner_radius=12)
        pane.add(left, minsize=260)

        self._style_tree()
        self._tree = ttk.Treeview(left, show="tree headings",
                                   columns=("count",), selectmode="browse")
        self._tree.heading("#0",    text="Klasör")
        self._tree.heading("count", text="Mail")
        self._tree.column("#0",    width=220)
        self._tree.column("count", width=60, anchor="center")
        self._tree.pack(fill="both", expand=True, padx=8, pady=8)
        self._tree.bind("<<TreeviewSelect>>", self._on_folder_select)

        # Right: mail list
        right = ctk.CTkFrame(pane, fg_color=CARD, corner_radius=12)
        pane.add(right, minsize=400)

        self._mail_list = ttk.Treeview(right,
            columns=("subject", "sender", "date"),
            show="headings", selectmode="browse")
        for col, lbl, w in [("subject","Konu",320),("sender","Gönderen",200),("date","Tarih",130)]:
            self._mail_list.heading(col, text=lbl)
            self._mail_list.column(col, width=w)
        self._mail_list.pack(fill="both", expand=True, padx=8, pady=8)
        self._mail_list.bind("<Double-1>", self._on_mail_double_click)

        # Bottom button row
        btn_row = ctk.CTkFrame(self._main, fg_color="transparent")
        btn_row.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(btn_row, text="Seçili Klasörü Tara & Kaydet",
                      height=40, fg_color=ACCENT, hover_color=AHOVER,
                      font=ctk.CTkFont(weight="bold"),
                      command=self._start_scan).pack(side="right")

        threading.Thread(target=self._load_tree, daemon=True).start()

    def _pst_summary(self) -> str:
        if not self._pst_files:
            return "Henüz PST dosyası eklenmedi"
        names = [os.path.basename(p) for p in self._pst_files]
        return "  |  ".join(names)

    def _add_pst_file(self):
        paths = filedialog.askopenfilenames(
            title="PST / OST Dosyası Seç",
            filetypes=[("PST / OST", "*.pst *.ost"), ("Tüm dosyalar", "*.*")]
        )
        if not paths:
            return
        
        for p in paths:
            if p not in self._pst_files:
                self._pst_files.append(p)
                # Fallback: If pypff is missing, try to mount via Outlook Engine
                if not is_pypff_available():
                    self._set_status(f"Outlook motoru ile bağlanıyor: {os.path.basename(p)}")
                    add_pst_store(p)

        try:
            self._pst_label.configure(text=self._pst_summary())
        except Exception:
            pass

        self._tree.delete(*self._tree.get_children())
        self._folder_map = {}
        
        if is_pypff_available():
            threading.Thread(target=self._load_tree_pst, daemon=True).start()
        else:
            # If no pypff, we use the COM-based tree loading even in PST mode
            # since add_pst_store just attached it to the COM session.
            threading.Thread(target=self._load_tree_com_logic, daemon=True).start()

    def _style_tree(self):
        s = ttk.Style()
        s.theme_use("default")
        s.configure("Treeview", background=CARD, foreground=TMAIN,
                    fieldbackground=CARD, borderwidth=0, rowheight=26)
        s.map("Treeview", background=[("selected", ACCENT)])
        s.configure("Treeview.Heading", background=SIDE, foreground=TDIM, borderwidth=0)

    def _load_tree(self):
        """Branch: COM mode."""
        if self._mode == "pst":
            if is_pypff_available():
                self._load_tree_pst()
            else:
                self._load_tree_com_logic()
            return
        self._load_tree_com_logic()

    def _load_tree_com_logic(self):
        self._set_status("Outlook klasörleri yükleniyor...")
        try:
            stores = enumerate_stores_tree()
        except Exception as e:
            self._set_status(f"Hata: {e}")
            return

        def update_ui():
            if not self._tree.winfo_exists():
                return
            self._tree.delete(*self._tree.get_children())
            self._folder_map = {}   # iid -> entry_id (COM) or dict (PST)

            def insert_node(parent_iid, node):
                iid = self._tree.insert(parent_iid, "end",
                                        text=node["name"],
                                        values=(node["item_count"],),
                                        open=(parent_iid == ""))
                self._folder_map[iid] = {"mode": "com", "entry_id": node["entry_id"]}
                for child in node.get("children", []):
                    insert_node(iid, child)

            for store in stores:
                store_iid = self._tree.insert("", "end",
                                              text=f"🗄  {store['store']}",
                                              values=("",), open=True)
                for child in store.get("children", []):
                    insert_node(store_iid, child)
            self._set_status(f"{len(stores)} store yüklendi.")

        self.after(0, update_ui)

    def _load_tree_pst(self):
        """Branch: PST file mode — reads each added PST via pypff."""
        if not self._pst_files:
            self._set_status("PST dosyası eklenmedi. '+ PST Ekle' butonunu kullanın.")
            return
        if not is_pypff_available():
            self.after(0, lambda: messagebox.showerror(
                "Eksik Kütüphane",
                "PST okuma için 'libpff-python' paketi gerekli.\n"
                "Yüklemek için: pip install libpff-python"))
            return

        pst_data = []
        for pst_path in self._pst_files:
            self._set_status(f"Okunuyor: {os.path.basename(pst_path)}")
            try:
                tree = read_pst_tree(pst_path)
                pst_data.append((pst_path, tree))
            except Exception as exc:
                self._set_status(f"Hata ({os.path.basename(pst_path)}): {exc}")
                continue

        def update_ui():
            if not self._tree.winfo_exists():
                return
            self._tree.delete(*self._tree.get_children())
            self._folder_map = {}

            def insert_pst_node(parent_iid, node):
                iid = self._tree.insert(parent_iid, "end",
                                        text=node["name"],
                                        values=(node["item_count"],),
                                        open=(parent_iid == ""))
                self._folder_map[iid] = {
                    "mode":       "pst",
                    "pst_path":   node["_pst_path"],
                    "index_path": node["_index_path"],
                }
                for child in node.get("children", []):
                    insert_pst_node(iid, child)

            for pst_path, tree in pst_data:
                store_iid = self._tree.insert("", "end",
                                              text=f"🗄  {os.path.basename(pst_path)}",
                                              values=("",), open=True)
                for child in tree.get("children", []):
                    insert_pst_node(store_iid, child)

            self._set_status(f"{len(self._pst_files)} PST yüklendi.")

        self.after(0, update_ui)

    def _on_folder_select(self, _event):
        sel = self._tree.selection()
        if not sel:
            return
        iid = sel[0]
        info = self._folder_map.get(iid)
        if not info:
            return
        self._mail_list.delete(*self._mail_list.get_children())
        self._set_status("Mailler yükleniyor...")
        threading.Thread(target=self._load_mails, args=(info,), daemon=True).start()

    def _load_mails(self, info: dict):
        if info["mode"] == "com":
            mails = extract_mails(info["entry_id"], limit=300)
        else:
            mails = extract_mails_from_pst(info["pst_path"], info["index_path"], limit=300)
        self.after(0, lambda: self._populate_mail_list(mails))

    def _populate_mail_list(self, mails):
        self._mail_list.delete(*self._mail_list.get_children())
        self._current_mails = mails
        for m in mails:
            dt = str(m.get("received_at", ""))[:16]
            self._mail_list.insert("", "end",
                values=(m.get("subject","")[:80], m.get("sender",""), dt))
        self._set_status(f"{len(mails)} mail gösteriliyor.")

    def _on_mail_double_click(self, _event):
        sel = self._mail_list.selection()
        if not sel:
            return
        idx = self._mail_list.index(sel[0])
        mails = getattr(self, "_current_mails", [])
        if idx < len(mails):
            self._show_mail_detail(mails[idx])

    def _show_mail_detail(self, mail):
        win = ctk.CTkToplevel(self)
        win.title(mail.get("subject", "Mail Detayı"))
        win.geometry("760x560")
        win.configure(fg_color=BG)

        info = (f"Konu:     {mail.get('subject','')}\n"
                f"Gönderen: {mail.get('sender','')}\n"
                f"Tarih:    {str(mail.get('received_at',''))[:19]}\n"
                f"Klasör:   {mail.get('folder_path','')}")
        ctk.CTkLabel(win, text=info, font=ctk.CTkFont(size=12),
                     text_color=TDIM, justify="left").pack(anchor="w", padx=20, pady=(16,8))
        ctk.CTkLabel(win, text="─"*80, text_color="#334155").pack(fill="x", padx=20)

        tb = ctk.CTkTextbox(win, fg_color=CARD, text_color=TMAIN, wrap="word")
        tb.pack(fill="both", expand=True, padx=20, pady=12)
        tb.insert("1.0", mail.get("body",""))
        tb.configure(state="disabled")

    def _start_scan(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Seçim", "Lütfen bir klasör seçin.")
            return
        iid = sel[0]
        info = self._folder_map.get(iid)
        if not info:
            messagebox.showwarning("Seçim", "Bu bir store köküdür, alt klasör seçin.")
            return
        self._set_status("Tarama başladı...")
        threading.Thread(target=self._run_scan, args=(info,), daemon=True).start()

    def _run_scan(self, info: dict):
        if info["mode"] == "com":
            mails = extract_mails(info["entry_id"], limit=500)
        else:
            mails = extract_mails_from_pst(info["pst_path"], info["index_path"], limit=500)
        db = SessionLocal()
        try:
            for i, m in enumerate(mails, 1):
                analysis = analyze_mail_content(m.get("body", ""))
                process_mail(db, {**m, **analysis})
                if i % 10 == 0:
                    self._set_status(f"Tarandı: {i}/{len(mails)}")
        finally:
            db.close()
        self.after(0, lambda: messagebox.showinfo("Tamamlandı", f"{len(mails)} mail kaydedildi."))
        self._set_status(f"Tarama tamamlandı: {len(mails)} mail.")

    # ──────────────────────────── SEARCH ────────────────────────────
    def show_search(self):
        self._clear_main()
        self._set_nav("search")

        ctk.CTkLabel(self._main, text="Global Arama",
                     font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", pady=(0,16))

        form = ctk.CTkFrame(self._main, fg_color=CARD, corner_radius=12)
        form.pack(fill="x")

        def field(parent, lbl):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", padx=32, pady=6)
            ctk.CTkLabel(row, text=lbl, width=140, anchor="w",
                         text_color=TDIM, font=ctk.CTkFont(size=12)).pack(side="left")
            e = ctk.CTkEntry(row, height=34, fg_color="#0f172a", border_color="#334155")
            e.pack(side="left", fill="x", expand=True)
            return e

        ctk.CTkLabel(form, text="", height=8).pack()
        self._s_subject  = field(form, "Konu içerir:")
        self._s_sender   = field(form, "Gönderen:")
        self._s_body     = field(form, "İçerik (anahtar kelime):")
        ctk.CTkLabel(form, text="", height=8).pack()

        ctk.CTkButton(form, text="  🔍  ARA  ", height=40,
                      fg_color=ACCENT, hover_color=AHOVER,
                      font=ctk.CTkFont(weight="bold", size=13),
                      command=self._run_search).pack(pady=(0,16))

        # Results area
        res_lbl = ctk.CTkLabel(self._main, text="Sonuçlar",
                               font=ctk.CTkFont(size=15, weight="bold"), anchor="w")
        res_lbl.pack(fill="x", pady=(20,6))
        self._s_count_lbl = ctk.CTkLabel(self._main, text="",
                                          font=ctk.CTkFont(size=12), text_color=TDIM, anchor="w")
        self._s_count_lbl.pack(fill="x")

        res_frame = ctk.CTkFrame(self._main, fg_color=CARD, corner_radius=12)
        res_frame.pack(fill="both", expand=True, pady=(6,0))

        self._search_tree = ttk.Treeview(res_frame,
            columns=("subject","sender","folder","date"),
            show="headings", selectmode="browse")
        for col, lbl, w in [("subject","Konu",300),("sender","Gönderen",180),
                             ("folder","Klasör",200),("date","Tarih",120)]:
            self._search_tree.heading(col, text=lbl)
            self._search_tree.column(col, width=w)
        sb2 = ttk.Scrollbar(res_frame, orient="vertical", command=self._search_tree.yview)
        self._search_tree.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y", pady=8)
        self._search_tree.pack(fill="both", expand=True, padx=8, pady=8)
        self._search_tree.bind("<Double-1>", self._on_search_double_click)
        self._search_results = []

    def _run_search(self):
        params = {
            "subject":       self._s_subject.get().strip(),
            "sender":        self._s_sender.get().strip(),
            "body_contains": self._s_body.get().strip(),
        }
        if not any(params.values()):
            messagebox.showwarning("Arama", "En az bir alan doldurun.")
            return
        self._set_status("Aranıyor...")
        self._s_count_lbl.configure(text="Aranıyor…")
        threading.Thread(target=self._search_worker, args=(params,), daemon=True).start()

    def _search_worker(self, params):
        if self._mode == "com":
            results = search_mails(params)
        else:
            # PST mode: search all added PST files
            if not self._pst_files:
                self.after(0, lambda: messagebox.showwarning(
                    "PST Yok", "Önce Explorer'dan PST dosyası ekleyin."))
                return
            results = []
            for p in self._pst_files:
                try:
                    results.extend(search_pst(p, params))
                except Exception as exc:
                    import logging
                    logging.getLogger(__name__).error("PST search error %s: %s", p, exc)
        self.after(0, lambda: self._populate_search(results))


    def _populate_search(self, results):
        self._search_results = results
        self._search_tree.delete(*self._search_tree.get_children())
        for m in results:
            dt = str(m.get("received_at",""))[:16]
            self._search_tree.insert("", "end",
                values=(m.get("subject","")[:80], m.get("sender",""),
                        m.get("folder_path","")[-50:], dt))
        total = len(results)
        senders = len({m.get("sender","") for m in results})
        self._s_count_lbl.configure(
            text=f"Toplam {total} sonuç  •  {senders} farklı gönderen")
        self._set_status(f"Arama tamamlandı: {total} sonuç.")

    def _on_search_double_click(self, _event):
        sel = self._search_tree.selection()
        if not sel:
            return
        idx = self._search_tree.index(sel[0])
        if idx < len(self._search_results):
            self._show_mail_detail(self._search_results[idx])

    # ──────────────────────────── STATS ────────────────────────────
    def show_stats(self):
        self._clear_main()
        self._set_nav("stats")

        ctk.CTkLabel(self._main, text="İstatistikler",
                     font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", pady=(0,16))

        db = SessionLocal()
        total   = db.query(Mail).count()
        high    = db.query(Mail).filter(Mail.risk_score >= 50).count()
        stores  = db.query(Mail.pst_source).distinct().count()
        folders = db.query(Mail.folder_path).distinct().count()

        from sqlalchemy import func
        top_senders = (db.query(Mail.sender, func.count(Mail.id).label("cnt"))
                       .group_by(Mail.sender).order_by(func.count(Mail.id).desc())
                       .limit(10).all())
        db.close()

        # Stat cards
        cards = ctk.CTkFrame(self._main, fg_color="transparent")
        cards.pack(fill="x", pady=(0,20))
        cards.grid_columnconfigure((0,1,2,3), weight=1)

        def card(parent, col, title, val, color):
            f = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=12, height=110)
            f.grid(row=0, column=col, padx=8, sticky="nsew")
            f.grid_propagate(False)
            accent = ctk.CTkFrame(f, width=4, fg_color=color, corner_radius=2)
            accent.pack(side="left", fill="y", padx=(14,0), pady=14)
            inner = ctk.CTkFrame(f, fg_color="transparent")
            inner.pack(side="left", padx=16, pady=14)
            ctk.CTkLabel(inner, text=title.upper(), font=ctk.CTkFont(size=10, weight="bold"),
                         text_color=TDIM).pack(anchor="w")
            ctk.CTkLabel(inner, text=str(val), font=ctk.CTkFont(size=28, weight="bold"),
                         text_color=TMAIN).pack(anchor="w")

        card(cards, 0, "Toplam Mail",  total,   ACCENT)
        card(cards, 1, "Yüksek Risk",  high,    DANGER)
        card(cards, 2, "Store Sayısı", stores,  SUCCESS)
        card(cards, 3, "Klasör Sayısı",folders, WARN)

        # Top senders table
        ctk.CTkLabel(self._main, text="En Çok Gönderen (DB'deki taranmış mailler)",
                     font=ctk.CTkFont(size=15, weight="bold"), anchor="w").pack(fill="x", pady=(4,8))

        tbl_frame = ctk.CTkFrame(self._main, fg_color=CARD, corner_radius=12)
        tbl_frame.pack(fill="both", expand=True)

        tbl = ttk.Treeview(tbl_frame, columns=("sender","count"), show="headings")
        tbl.heading("sender", text="Gönderen")
        tbl.heading("count",  text="Mail Sayısı")
        tbl.column("sender", width=420)
        tbl.column("count",  width=120, anchor="center")
        tbl.pack(fill="both", expand=True, padx=10, pady=10)

        for sender, cnt in top_senders:
            tbl.insert("", "end", values=(sender, cnt))

    # ──────────────────────────── RESULTS ────────────────────────────
    def show_results(self):
        self._clear_main()
        self._set_nav("results")

        ctk.CTkLabel(self._main, text="Analiz & Tespit Logu",
                     font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", pady=(0,12))

        db = SessionLocal()
        mails = db.query(Mail).order_by(Mail.risk_score.desc()).limit(200).all()
        db.close()

        frame = ctk.CTkFrame(self._main, fg_color=CARD, corner_radius=12)
        frame.pack(fill="both", expand=True)

        cols = ("score","subject","sender","folder","source")
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for col, lbl, w in [("score","Risk",60),("subject","Konu",280),
                             ("sender","Gönderen",180),("folder","Klasör",200),
                             ("source","Store",130)]:
            tree.heading(col, text=lbl)
            tree.column(col, width=w, anchor="center" if col=="score" else "w")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", pady=8)
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        for m in mails:
            tree.insert("", "end", values=(
                f"{m.risk_score}%",
                (m.subject or "")[:60],
                m.sender or "",
                (m.folder_path or "")[-50:],
                m.pst_source or ""))

        ctk.CTkButton(self._main, text="📥 XLSX Rapor Al",
                      fg_color=SUCCESS, hover_color="#059669",
                      font=ctk.CTkFont(weight="bold"),
                      command=self._export).pack(side="right", pady=(12,0))

    def _export(self):
        from backend.app.services.report_service import generate_excel_report
        db = SessionLocal()
        path = generate_excel_report(db)
        db.close()
        if path:
            messagebox.showinfo("Export", f"Kaydedildi:\n{path}")


if __name__ == "__main__":
    App().mainloop()
