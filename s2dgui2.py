import customtkinter as ctk
import threading
import time
import datetime
import os
import string

# ==========================================
# AYARLAR & YARDIMCILAR
# ==========================================
TEST_MODE = True  # True: Test Modu (Excel gerekmez) / False: Gerçek Mod

def col2num(col_str):
    """Harfi sayıya çevirir (A->1, Z->26, AA->27)"""
    num = 0
    col_str = col_str.strip().upper()
    for c in col_str:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c) - ord('A')) + 1
    return num

def num2col(n):
    """Sayıyı harfe çevirir (1->A, 27->AA)"""
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

# ==========================================
# ARKA PLAN İŞÇİSİ (DATA OKUMA & YAZMA)
# ==========================================
class WorkerThread(threading.Thread):
    def __init__(self, app, excel_path, config, dynamic_params):
        super().__init__()
        self.app = app
        self.excel_path = excel_path
        self.config = config
        self.dynamic_params = dynamic_params # List of tuples: (Suffix, ColChar)
        self.running = True
        self.daemon = True

    def run(self):
        if TEST_MODE:
            self.run_simulation()
        else:
            self.run_real_process()

    def run_simulation(self):
        sheet_name = self.config.get("sheet_name", "Sheet1")
        self.app.log(f"Başladı: {sheet_name}", "info")
        
        # Dinamik parametreleri logla
        msg = "Parametreler: "
        for suffix, col in self.dynamic_params:
            msg += f"[{suffix}->{col}] "
        self.app.log(msg, "info")
        
        time.sleep(1)
        total = 50
        self.app.after(0, self.app.update_max_progress, total)
        
        updates = 0
        errors = 0
        
        for i in range(1, total + 1):
            if not self.running: break
            time.sleep(0.05)
            self.app.after(0, self.app.update_stats, i, updates, errors)
            
            if i % 5 == 0:
                updates += 1
                # Örnek: İlk dinamik parametreyi update etmiş gibi yapalım
                p_name = f"Rib_{i}_{self.dynamic_params[0][0]}" if self.dynamic_params else "Param"
                self.app.after(0, self.app.log, f"{p_name} güncellendi.", "update")
            
        self.app.after(0, self.app.finish_process)

    def run_real_process(self):
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()

            # catia = win32com.client.GetActiveObject("CATIA.Application")
            excel = win32com.client.Dispatch("Excel.Application")
            
            wb = excel.Workbooks.Open(self.excel_path)
            target_sheet = self.config.get("sheet_name", "")
            
            # Sayfa bul
            try: valid_sheet = wb.Sheets(target_sheet)
            except: valid_sheet = wb.ActiveSheet

            last_row = valid_sheet.Cells(valid_sheet.Rows.Count, 1).End(-4162).Row
            
            # Excel verisini komple belleğe al (Hız için)
            # Sütun sınırını bulalım (En az Z sütununa kadar okusun)
            raw_data = valid_sheet.Range(f"A2:Z{last_row}").Value 
            # raw_data tuple of tuples döner. raw_data[satir_idx][sutun_idx]
            
            total_rows = len(raw_data)
            self.app.after(0, self.app.update_max_progress, total_rows)
            
            updates = 0
            errors = 0
            
            # --- PARAMETRE MAP ---
            # Hangi sütun (index) hangi Suffix'e gidecek?
            # Örn: [('Thickness', 1), ('H', 2), ('P1', 10)]  (0-based index for python list)
            param_map = []
            for suffix, col_letter in self.dynamic_params:
                if not col_letter: continue
                col_idx = col2num(col_letter) - 1 # Python 0-based index
                param_map.append((suffix, col_idx))
            
            # DÖNGÜ
            for i, row in enumerate(raw_data):
                if not self.running: break
                
                # ID Okuma (A Sütunu -> index 0)
                try: id_val = row[0]
                except: continue
                
                if not id_val: continue
                
                # ID Formatlama (101.0 -> "101")
                try: id_str = str(int(float(id_val))) if isinstance(id_val, float) and id_val.is_integer() else str(id_val).strip()
                except: id_str = str(id_val).strip()
                
                # Dinamik Parametreleri Güncelle
                for suffix, col_idx in param_map:
                    if col_idx < len(row):
                        val = row[col_idx]
                        if val is not None:
                            # CATIA Parametre Adı: ID + Suffix (Örn: Rib101 + P1 -> Rib101P1)
                            # Eğer Suffix boşsa direkt ID (Thickness gibi)
                            full_name = id_str + suffix
                            
                            # --- CATIA YAZMA KODU BURAYA GELECEK ---
                            # p = part.Parameters.Item(full_name)
                            # p.Value = float(val)
                            # updates += 1
                            pass

                self.app.after(0, self.app.update_stats, i+1, updates, errors)

            wb.Close(False)
            excel.Quit()
            self.app.after(0, self.app.finish_process)

        except Exception as e:
            self.app.after(0, self.app.log, f"KRİTİK HATA: {e}", "error")

    def stop(self):
        self.running = False

# ==========================================
# EXCEL PREVIEW & ANALİZ
# ==========================================
class ExcelPreviewLoader(threading.Thread):
    def __init__(self, app, path):
        super().__init__()
        self.app = app
        self.path = path
        self.daemon = True

    def run(self):
        try:
            data_preview = []
            sheets = []
            
            if TEST_MODE:
                time.sleep(0.5)
                sheets = ["Visualization Data", "Sheet2"]
                # Fake Data: Header + 5 Rows
                data_preview = [
                    ["ID", "Th", "Height", "Mat", "X", "Y", "Z", "Q", "W", "E", "P1", "D1", "P2", "D2"],
                    ["101", "5.0", "20.0", "AL", "10", "20", "30", "-", "-", "-", "1.5", "0.5", "2.0", "0.8"],
                    ["102", "5.2", "21.0", "AL", "12", "22", "32", "-", "-", "-", "1.6", "0.6", "2.1", "0.9"],
                    ["103", "5.4", "22.0", "ST", "14", "24", "34", "-", "-", "-", "1.7", "0.7", "2.2", "1.0"],
                    ["104", "5.6", "23.0", "ST", "16", "26", "36", "-", "-", "-", "1.8", "0.8", "2.3", "1.1"],
                ]
            else:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(self.path, ReadOnly=True)
                
                # Sayfaları al
                for s in wb.Sheets:
                    sheets.append(s.Name)
                
                # İlk sayfadan önizleme al (İlk 10 satır, ilk 15 sütun)
                ws = wb.Sheets(1)
                vals = ws.Range("A1:O10").Value
                # Tuple to List
                if vals:
                    for row in vals:
                        data_preview.append(list(row) if row else [])
                
                wb.Close(False)
                excel.Quit()
            
            self.app.after(0, self.app.update_ui_with_excel_data, sheets, data_preview)

        except Exception as e:
            self.app.after(0, self.app.log, f"Önizleme hatası: {e}", "error")

# ==========================================
# ARAYÜZ (GUI)
# ==========================================
class AutomationSuite(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Config
        self.config = {"sheet_name": "", "always_on_top": False}
        self.param_rows = [] 

        # Tema ve Pencere
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("dark-blue") # Mavi vurgular
        
        # Renk Paleti (Custom Colors)
        self.colors = {
            "bg": "#1a1a1a",            # Ana arka plan (Daha koyu)
            "panel": "#2b2b2b",         # Paneller
            "card": "#333333",          # Parametre kartları
            "accent": "#3B8ED0",        # Mavi vurgu
            "danger": "#cf6679",        # Kırmızı (Silme)
            "text_gray": "#a1a1aa",     # Gri yazı
            "grid_header": "#404040",   # Tablo başlığı
            "grid_row_even": "#262626", # Tablo satır 1
            "grid_row_odd": "#1f1f1f"   # Tablo satır 2
        }
        self.title("CATIA Automation Suite v4.5 Pro")
        self.geometry("1100x750")

        # --- TAB YAPISI ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tab_view = ctk.CTkTabview(self, corner_radius=15, fg_color=self.colors["bg"])
        self.tab_view.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_monitor = self.tab_view.add("  Monitör  ")
        self.tab_settings = self.tab_view.add("  Ayarlar & Önizleme  ")
        self.setup_monitor()
        self.setup_settings()
        
        # Değişkenler
        self.selected_file = None
        self.worker = None
        self.total_work = 1
        self.start_time = 0

    # ------------------------------
    # MONİTÖR TABI (Aynı kalabilir veya ufak makyaj)
    # ------------------------------
    def setup_monitor(self):
        self.tab_monitor.grid_columnconfigure(0, weight=1)
        self.tab_monitor.grid_rowconfigure(2, weight=1)
        # Üst Panel
        frame_top = ctk.CTkFrame(self.tab_monitor, fg_color="transparent")
        frame_top.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        self.btn_file = ctk.CTkButton(frame_top, text="📂 Excel Dosyası Seç", command=self.select_file, 
                                      width=160, height=40, font=("Roboto", 13, "bold"))
        self.btn_file.pack(side="left", padx=5)
        
        self.lbl_file = ctk.CTkLabel(frame_top, text="Henüz dosya seçilmedi", text_color=self.colors["text_gray"], font=("Roboto", 12))
        self.lbl_file.pack(side="left", padx=15)
        
        self.btn_run = ctk.CTkButton(frame_top, text="İŞLEMİ BAŞLAT ▶", state="disabled", 
                                     command=self.start_process, fg_color="#2ECC71", 
                                     width=160, height=40, font=("Roboto", 13, "bold"))
        self.btn_run.pack(side="right", padx=5)
        # İstatistik Kartları
        frame_stats = ctk.CTkFrame(self.tab_monitor, fg_color=self.colors["panel"], corner_radius=10)
        frame_stats.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        frame_stats.columnconfigure((0,1,2,3), weight=1)
        
        self.lbl_time = self._card(frame_stats, 0, "Geçen Süre")
        self.lbl_eta = self._card(frame_stats, 1, "Tahmini Bitiş")
        self.lbl_upd = self._card(frame_stats, 2, "Başarılı", "#2ECC71")
        self.lbl_err = self._card(frame_stats, 3, "Hatalar", "#cf6679")
        # Log
        self.log_box = ctk.CTkTextbox(self.tab_monitor, font=("Consolas", 12), state="disabled", 
                                      fg_color="#111", text_color="#eee", corner_radius=10)
        self.log_box.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        # Progress
        self.progress = ctk.CTkProgressBar(self.tab_monitor, height=15, corner_radius=8)
        self.progress.grid(row=3, column=0, padx=10, pady=(5,15), sticky="ew")
        self.progress.set(0)
        
    def _card(self, parent, col, title, color="white"):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=0, column=col, padx=10, pady=10)
        ctk.CTkLabel(f, text=title, font=("Roboto", 11), text_color=self.colors["text_gray"]).pack()
        l = ctk.CTkLabel(f, text="--", font=("Roboto", 20, "bold"), text_color=color)
        l.pack()
        return l

    # ------------------------------
    # AYARLAR & ÖNİZLEME TABI (YENİ TASARIM)
    # ------------------------------
    def setup_settings(self):
        # Grid: Sol (4 birim), Sağ (6 birim)
        self.tab_settings.grid_columnconfigure(0, weight=2) # Sol Panel
        self.tab_settings.grid_columnconfigure(1, weight=3) # Sağ Panel
        self.tab_settings.grid_rowconfigure(0, weight=1)
        # --- SOL PANEL: PARAMETRELER ---
        left_container = ctk.CTkFrame(self.tab_settings, fg_color="transparent")
        left_container.grid(row=0, column=0, padx=(0, 10), pady=10, sticky="nsew")
        
        # Başlık Alanı
        header_lbl = ctk.CTkLabel(left_container, text="Parametre Eşleştirme", 
                                 font=("Roboto", 18, "bold"), text_color="white", anchor="w")
        header_lbl.pack(fill="x", pady=(0, 15))
        # Scrollable Parametre Listesi
        self.scroll_params = ctk.CTkScrollableFrame(left_container, label_text="", fg_color="transparent")
        self.scroll_params.pack(fill="both", expand=True, pady=(0, 10))
        # Varsayılan Parametreler
        defaults = [
            ("Thickness", "B"), ("H", "C"),
            ("P1", "K"), ("D1", "L"),
            ("P2", "M"), ("D2", "N")
        ]
        for name, col in defaults:
            self.add_param_row(name, col)
        # Ekleme Butonu
        btn_add = ctk.CTkButton(left_container, text="+ Parametre Ekle", 
                               command=lambda: self.add_param_row("", ""), 
                               fg_color=self.colors["panel"], hover_color=self.colors["card"],
                               text_color=self.colors["accent"], font=("Roboto", 12, "bold"),
                               height=40, border_width=1, border_color=self.colors["card"])
        btn_add.pack(fill="x", pady=(0, 15))

        # Alt Ayarlar (Sayfa Seçimi)
        settings_card = ctk.CTkFrame(left_container, fg_color=self.colors["panel"], corner_radius=10)
        settings_card.pack(fill="x")
        
        ctk.CTkLabel(settings_card, text="Hedef Excel Sayfası", font=("Roboto", 12, "bold"), text_color=self.colors["text_gray"]).pack(anchor="w", padx=15, pady=(10,5))
        self.combo_sheet = ctk.CTkComboBox(settings_card, values=["Dosya Bekleniyor..."], height=35, font=("Roboto", 13))
        self.combo_sheet.pack(fill="x", padx=15, pady=(0, 15))
        # --- SAĞ PANEL: CANLI ÖNİZLEME ---
        right_container = ctk.CTkFrame(self.tab_settings, fg_color=self.colors["panel"], corner_radius=15)
        right_container.grid(row=0, column=1, padx=(10, 0), pady=10, sticky="nsew")
        
        # Başlık
        ctk.CTkLabel(right_container, text="Canlı Veri Önizleme", font=("Roboto", 16, "bold")).pack(pady=15)
        
        # Tablo Alanı (Scrollable)
        self.preview_box = ctk.CTkScrollableFrame(right_container, fg_color="#1e1e1e", corner_radius=10)
        self.preview_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        # Empty State
        self.empty_state_frame = ctk.CTkFrame(self.preview_box, fg_color="transparent")
        self.empty_state_frame.pack(expand=True, fill="both", pady=100)
        ctk.CTkLabel(self.empty_state_frame, text="📊", font=("Arial", 40)).pack()
        ctk.CTkLabel(self.empty_state_frame, text="Önizleme için bir Excel dosyası seçin.", 
                     font=("Roboto", 14), text_color="gray").pack(pady=10)

    def add_param_row(self, default_name, default_col):
        """
        Daha şık, 'Kart' görünümlü parametre satırı.
        """
        # Kart Konteyner
        card = ctk.CTkFrame(self.scroll_params, fg_color=self.colors["card"], corner_radius=8)
        card.pack(fill="x", pady=4, padx=2)
        
        # İçerik için bir frame (pack kullanarak)
        content_frame = ctk.CTkFrame(card, fg_color="transparent")
        content_frame.pack(fill="x", padx=10, pady=10)
        
        # 1. Sütun Harfi (Badge Style)
        col_container = ctk.CTkFrame(content_frame, fg_color="transparent", width=60)
        col_container.pack(side="left", padx=(0, 5))
        
        ctk.CTkLabel(col_container, text="Sütun", font=("Arial", 10), text_color="gray").pack(anchor="w")
        col_entry = ctk.CTkEntry(col_container, width=50, height=30, 
                                 font=("Roboto", 14, "bold"), justify="center",
                                 fg_color=self.colors["bg"], border_width=0)
        col_entry.insert(0, default_col)
        col_entry.pack()
        
        # Ok İkonu
        ctk.CTkLabel(content_frame, text="➜", text_color="gray", font=("Arial", 14)).pack(side="left", padx=5)
        
        # 2. Parametre Adı (Suffix)
        name_container = ctk.CTkFrame(content_frame, fg_color="transparent")
        name_container.pack(side="left", fill="x", expand=True, padx=5)
        
        ctk.CTkLabel(name_container, text="CATIA Parametre Adı (Ek)", font=("Arial", 10), text_color="gray").pack(anchor="w")
        name_entry = ctk.CTkEntry(name_container, height=30, font=("Roboto", 13),
                                  placeholder_text="Örn: Thickness", 
                                  fg_color=self.colors["bg"], border_width=0)
        name_entry.insert(0, default_name)
        name_entry.pack(fill="x")
        
        # 3. Canlı ID Önizlemesi (Küçük bilgi)
        info_lbl = ctk.CTkLabel(content_frame, text=f"ID{default_name}", font=("Consolas", 11), text_color=self.colors["accent"])
        info_lbl.pack(side="left", padx=10)
        
        # 4. Silme Butonu (Ghost Style)
        del_btn = ctk.CTkButton(content_frame, text="✕", width=30, height=30,
                                fg_color="transparent", text_color="gray",
                                hover_color=self.colors["danger"],
                                font=("Arial", 14, "bold"),
                                command=lambda: self.delete_param_row(card))
        del_btn.pack(side="right", padx=5)
        
        # Referansları Sakla
        self.param_rows.append({
            "frame": card,
            "col": col_entry,
            "name": name_entry,
            "info": info_lbl
        })
        
        # Dinamik Yazı Güncelleme
        def update_info(event=None):
            suf = name_entry.get()
            info_lbl.configure(text=f"ID{suf}" if suf else "ID...")
        
        name_entry.bind("<KeyRelease>", update_info)

    def delete_param_row(self, frame):
        frame.destroy()
        # Listeden temizle
        self.param_rows = [r for r in self.param_rows if r["frame"].winfo_exists()]

    # ------------------------------
    # İŞLEVLER
    # ------------------------------
    def select_file(self):
        self.attributes("-topmost", False)
        path = ctk.filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm")])
        
        if path:
            self.selected_file = path
            self.lbl_file.configure(text=os.path.basename(path), text_color="white")
            self.btn_run.configure(state="normal")
            self.log(f"Dosya seçildi: {path}", "info")
            
            # Tabı Ayarlar'a geçir
            self.tab_view.set("  Ayarlar & Önizleme  ")
            self.log("Önizleme oluşturuluyor...", "info")
            ExcelPreviewLoader(self, path).start()

    def update_ui_with_excel_data(self, sheets, data):
        # Sayfa listesini güncelle
        self.combo_sheet.configure(values=sheets)
        if sheets: self.combo_sheet.set(sheets[0])
        
        # Eski widgetları temizle
        for widget in self.preview_box.winfo_children():
            widget.destroy()
            
        if not data:
            ctk.CTkLabel(self.preview_box, text="Veri okunamadı.", text_color="red").pack(pady=20)
            return
        # --- YENİ TABLO TASARIMI ---
        
        # Sütun Genişliklerini Ayarla
        num_cols = len(data[0]) if data else 0
        for i in range(num_cols + 1): # +1 satır numarası için
            self.preview_box.grid_columnconfigure(i, weight=1)
        # 1. Header Satırı (A, B, C...)
        # 0. sütun boş (Satır no için)
        for c_idx in range(num_cols):
            col_letter = num2col(c_idx + 1)
            header = ctk.CTkLabel(self.preview_box, text=col_letter, 
                                  fg_color=self.colors["grid_header"], 
                                  font=("Roboto", 12, "bold"), height=30, corner_radius=4)
            header.grid(row=0, column=c_idx+1, padx=1, pady=(0,2), sticky="ew")
        # 2. Veri Satırları
        for r_idx, row_data in enumerate(data):
            # Satır Numarası (Sol Kenar)
            row_num = ctk.CTkLabel(self.preview_box, text=str(r_idx+1), width=30, 
                                   text_color="gray", font=("Arial", 10))
            row_num.grid(row=r_idx+1, column=0, padx=2)
            
            bg_color = self.colors["grid_row_even"] if r_idx % 2 == 0 else self.colors["grid_row_odd"]
            
            for c_idx, cell_val in enumerate(row_data):
                val_str = str(cell_val) if cell_val is not None else ""
                # Uzun metinleri kes
                if len(val_str) > 12: val_str = val_str[:9] + "..."
                
                cell = ctk.CTkLabel(self.preview_box, text=val_str, height=28,
                                    fg_color=bg_color, font=("Consolas", 11))
                cell.grid(row=r_idx+1, column=c_idx+1, padx=1, pady=1, sticky="ew")
        self.log("Önizleme yüklendi.", "success")

    def start_process(self):
        dynamic_params = []
        for row in self.param_rows:
            suffix = row["name"].get().strip()
            col = row["col"].get().strip()
            if col: dynamic_params.append((suffix, col))
        
        self.config["sheet_name"] = self.combo_sheet.get()
        
        self.tab_view.set("  Monitör  ")
        self.btn_run.configure(state="disabled", text="ÇALIŞIYOR...")
        self.log_box.configure(state="normal"); self.log_box.delete("1.0", "end"); self.log_box.configure(state="disabled")
        
        self.start_time = time.time()
        self.worker = WorkerThread(self, self.selected_file, self.config, dynamic_params)
        self.worker.start()

    # (Diğer yardımcı fonksiyonlar: log, update_stats vs. aynı kalır)
    def log(self, msg, type="info"):
        icons = {"info": "ℹ", "update": "⚡", "error": "✖", "success": "✔"}
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {icons.get(type, '')} {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        
    def update_stats(self, done, updates, errors):
        ratio = done / self.total_work if self.total_work > 0 else 0
        self.progress.set(ratio)
        self.lbl_upd.configure(text=str(updates))
        self.lbl_err.configure(text=str(errors))
        elapsed = time.time() - self.start_time
        if elapsed > 0:
            speed = done/elapsed if elapsed > 0 else 0.001
            rem = (self.total_work - done) / speed
            self.lbl_time.configure(text=time.strftime('%M:%S', time.gmtime(elapsed)))
            self.lbl_eta.configure(text=time.strftime('%M:%S', time.gmtime(rem)))

    def update_max_progress(self, val): 
        self.total_work = val
        self.start_time = time.time()
        self.progress.set(0)
        
    def finish_process(self):
        self.btn_run.configure(state="normal", text="İŞLEMİ BAŞLAT ▶")
        self.log("İşlem tamamlandı.", "success")
        
    def stop_process(self):
        if self.worker: self.worker.stop()

if __name__ == "__main__":
    app = AutomationSuite()
    app.mainloop()