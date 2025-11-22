import customtkinter as ctk
import threading
import time
import datetime
import os

# ==========================================
# AYARLAR
# ==========================================
TEST_MODE = True  # True: Test Modu / False: Gerçek Mod
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# ==========================================
# İŞ PARÇACIĞI (WORKER THREAD)
# ==========================================
class WorkerThread(threading.Thread):
    """
    Tüm ağır işleri yapan arka plan işçisi.
    Arayüzü dondurmadan çalışır.
    """
    def __init__(self, app, excel_path, settings):
        super().__init__()
        self.app = app
        self.excel_path = excel_path
        self.settings = settings
        self.running = True
        self.daemon = True # Ana program kapanınca bu da ölsün

    def run(self):
        # --- TEST MODU SİMÜLASYONU ---
        if TEST_MODE:
            self.run_simulation()
        else:
            # --- GERÇEK CATIA KODU BURAYA ---
            self.run_real_process()

    def run_simulation(self):
        total_items = 200
        self.app.update_max_progress(total_items)
        
        updates = 0
        errors = 0
        
        self.app.log("Test simülasyonu başlatıldı...", "info")
        time.sleep(1) # Bağlantı simülasyonu

        for i in range(1, total_items + 1):
            if not self.running: break
            
            time.sleep(0.03) # İşlem süresi
            
            # UI Güncelleme (Callback ile)
            self.app.after(0, self.app.update_stats, i, updates, errors)
            
            # Senaryolar
            if i % 10 == 0:
                updates += 1
                self.app.after(0, self.app.log, f"Parametre_{i} güncellendi -> {i*2.5}", "update")
            
            if i == 55:
                errors += 1
                self.app.after(0, self.app.log, f"ID_{i} Excel'de hatalı format!", "error")

        if self.running:
            self.app.after(0, self.app.finish_process)
        else:
            self.app.after(0, self.app.log, "İşlem Thread içinde durduruldu.", "error")

    def run_real_process(self):
        # BURASI SENİN ORİJİNAL KOD MANTIĞIN
        # Ancak UI güncellemelerini 'self.app.after(0, ...)' ile yapmalısın.
        # Çünkü Tkinter thread-safe değildir.
        try:
            import win32com.client
            pythoncom = __import__("pythoncom") # Thread içinde COM kullanmak için şart
            pythoncom.CoInitialize() # COM başlat

            catia = win32com.client.GetActiveObject("CATIA.Application")
            excel = win32com.client.Dispatch("Excel.Application")
            
            # Excel Açma
            wb = excel.Workbooks.Open(self.excel_path)
            sheet = wb.Sheets("Visualization Data")
            last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
            
            # Veriyi okuma
            # ... (Veri okuma kodların buraya) ... 
            # Örnek:
            total_work = 100 # Gerçek sayıyı hesapla
            self.app.after(0, self.app.update_max_progress, total_work)

            # Döngü
            # for ...
            #    if not self.running: break
            #    ... işlemleri yap ...
            #    self.app.after(0, self.app.update_stats, current, u, e)
            
            wb.Close(False)
            excel.Quit()
            # catia.ActiveEditor.ActiveObject.Update()
            
            self.app.after(0, self.app.finish_process)

        except Exception as e:
            self.app.after(0, self.app.log, f"KRİTİK HATA: {e}", "error")

    def stop(self):
        self.running = False

# ==========================================
# ANA ARAYÜZ (DASHBOARD)
# ==========================================
class ProDashboard(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Durum Değişkenleri
        self.worker = None
        self.total_work = 1
        self.start_time = 0
        self.selected_file = None

        # Pencere
        self.title("CATIA Parameter Automation Pro")
        self.geometry("700x650")
        self.attributes("-topmost", True) 
        
        # Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1) # Log alanı esnek

        # --- 1. HEADER ---
        self.header = ctk.CTkFrame(self, fg_color="transparent")
        self.header.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        
        ctk.CTkLabel(self.header, text="Sizing Optimization Tool", font=("Roboto Medium", 24)).pack(side="left")
        
        self.status_badge = ctk.CTkLabel(
            self.header, text="BEKLENİYOR", text_color="gray", font=("Roboto", 12, "bold")
        )
        self.status_badge.pack(side="right", padx=10)

        # --- 2. DOSYA SEÇİM ALANI (Kontrol Paneli) ---
        self.control_panel = ctk.CTkFrame(self, fg_color="#2B2B2B", corner_radius=10)
        self.control_panel.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="ew")
        
        self.btn_select_file = ctk.CTkButton(
            self.control_panel, text="Excel Dosyası Seç", command=self.select_file, 
            fg_color="#3B8ED0", font=("Roboto", 12, "bold")
        )
        self.btn_select_file.pack(side="left", padx=15, pady=15)
        
        self.lbl_file_info = ctk.CTkLabel(self.control_panel, text="Dosya seçilmedi", text_color="gray")
        self.lbl_file_info.pack(side="left", padx=10)

        self.btn_start = ctk.CTkButton(
            self.control_panel, text="BAŞLAT", state="disabled", command=self.start_process,
            fg_color="#2ECC71", hover_color="#25a25a", width=100, font=("Roboto", 12, "bold")
        )
        self.btn_start.pack(side="right", padx=15)

        # --- 3. İSTATİSTİKLER ---
        self.stats_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.stats_frame.grid(row=2, column=0, padx=20, pady=0, sticky="ew")
        self.stats_frame.grid_columnconfigure((0,1,2,3), weight=1)

        self.card_elapsed = self.create_card(0, "Süre")
        self.card_eta = self.create_card(1, "Tahmini")
        self.card_updates = self.create_card(2, "Güncelleme", "#2ECC71")
        self.card_errors = self.create_card(3, "Hata", "#E74C3C")

        # --- 4. PROGRESS BAR ---
        self.progress_bar = ctk.CTkProgressBar(self, height=10)
        self.progress_bar.grid(row=4, column=0, padx=20, pady=(20, 5), sticky="ew")
        self.progress_bar.set(0)
        
        self.lbl_percent = ctk.CTkLabel(self, text="%0", font=("Roboto", 12))
        self.lbl_percent.grid(row=5, column=0, padx=20, sticky="e")

        # --- 5. LOG ALANI ---
        self.log_box = ctk.CTkTextbox(self, font=("Consolas", 11), activate_scrollbars=True)
        self.log_box.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="nsew")
        self.log("Sistem hazır. Lütfen bir Excel dosyası seçin.", "info")

        # --- 6. ACTION BAR (Alt Butonlar) ---
        self.action_bar = ctk.CTkFrame(self, fg_color="transparent")
        self.action_bar.grid(row=6, column=0, padx=20, pady=20, sticky="ew")
        
        self.btn_save_log = ctk.CTkButton(
            self.action_bar, text="Logu Kaydet", command=self.save_log, state="disabled",
            fg_color="#555", hover_color="#666"
        )
        self.btn_save_log.pack(side="left")
        
        self.btn_stop = ctk.CTkButton(
            self.action_bar, text="ÇIKIŞ", command=self.close_app,
            fg_color="#C0392B", hover_color="#E74C3C"
        )
        self.btn_stop.pack(side="right")

    def create_card(self, col, title, color="white"):
        f = ctk.CTkFrame(self.stats_frame, fg_color="#232323", corner_radius=10)
        f.grid(row=0, column=col, padx=5, pady=10, sticky="ew")
        ctk.CTkLabel(f, text=title, font=("Roboto", 11), text_color="gray").pack(pady=(5,0))
        lbl_val = ctk.CTkLabel(f, text="--", font=("Roboto", 18, "bold"), text_color=color)
        lbl_val.pack(pady=(0,5))
        return lbl_val

    # --- FONKSİYONLAR ---
    
    def select_file(self):
        file_path = ctk.filedialog.askopenfilename(
            title="Excel Seç", 
            filetypes=[("Excel Files", "*.xlsx *.xlsm")]
        )
        if file_path:
            self.selected_file = file_path
            filename = os.path.basename(file_path)
            size = round(os.path.getsize(file_path) / 1024, 1) # KB
            self.lbl_file_info.configure(text=f"{filename} ({size} KB)", text_color="white")
            self.btn_start.configure(state="normal")
            self.log(f"Dosya seçildi: {filename}", "info")

    def start_process(self):
        if not self.selected_file: return
        
        # UI Hazırlığı
        self.btn_start.configure(state="disabled", text="ÇALIŞIYOR...")
        self.btn_select_file.configure(state="disabled")
        self.btn_stop.configure(text="DURDUR")
        self.status_badge.configure(text="İŞLENİYOR", text_color="#3B8ED0")
        self.log_box.delete("1.0", "end") # Logu temizle
        
        self.start_time = time.time()
        
        # Thread Başlatma
        self.worker = WorkerThread(self, self.selected_file, {})
        self.worker.start()

    def update_max_progress(self, total):
        self.total_work = total

    def update_stats(self, done, updates, errors):
        # Thread'den çağrılır, güvenli UI güncellemesi
        ratio = done / self.total_work
        self.progress_bar.set(ratio)
        self.lbl_percent.configure(text=f"%{int(ratio*100)}")
        
        self.card_updates.configure(text=str(updates))
        self.card_errors.configure(text=str(errors))
        
        elapsed = time.time() - self.start_time
        speed = done / elapsed if elapsed > 0 else 0
        remaining = (self.total_work - done) / speed if speed > 0 else 0
        
        self.card_elapsed.configure(text=time.strftime('%M:%S', time.gmtime(elapsed)))
        self.card_eta.configure(text=time.strftime('%M:%S', time.gmtime(remaining)))

    def log(self, msg, type="info"):
        icons = {"info": "•", "update": "⚡", "error": "✖", "success": "✔"}
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{ts}] {icons.get(type, '•')} {msg}\n")
        self.log_box.see("end")

    def finish_process(self):
        self.status_badge.configure(text="TAMAMLANDI", text_color="#2ECC71")
        self.btn_start.configure(text="BAŞLAT") # Yeniden başlatmaya izin ver
        self.btn_select_file.configure(state="normal")
        self.btn_stop.configure(text="KAPAT", fg_color="#2B2B2B")
        self.btn_save_log.configure(state="normal", fg_color="#3B8ED0")
        self.log("İşlem başarıyla tamamlandı.", "success")
        
        # Kullanıcıya sesli/görsel uyarı (Windows top)
        self.attributes("-topmost", False)
        self.attributes("-topmost", True)

    def save_log(self):
        try:
            with open("Sizing_Log.txt", "w", encoding="utf-8") as f:
                f.write(self.log_box.get("1.0", "end"))
            self.log("Log dosyası 'Sizing_Log.txt' olarak kaydedildi.", "success")
        except Exception as e:
            self.log(f"Log kaydedilemedi: {e}", "error")

    def close_app(self):
        if self.worker and self.worker.is_alive():
            self.worker.stop() # Thread'i durdur
            self.log("İşlem durduruluyor...", "error")
            # Thread'in durması için biraz bekle
            self.after(500, self.destroy)
        else:
            self.destroy()

if __name__ == "__main__":
    app = ProDashboard()
    app.mainloop()