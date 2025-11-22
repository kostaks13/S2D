# 🚀 CATIA Automation Suite v4.5 Pro

Modern ve kullanıcı dostu CATIA otomasyon aracı - Excel verilerini CATIA parametrelerine otomatik aktarır.

## ✨ Özellikler

### 🎯 Temel Özellikler
- ✅ **Excel → CATIA Entegrasyonu**: Excel'deki verileri otomatik olarak CATIA parametrelerine yazar
- ✅ **Canlı Önizleme**: Excel dosyasının içeriğini anında görüntüleyin
- ✅ **Dinamik Parametre Eşleştirme**: İstediğiniz kadar parametre ekleyin/silin
- ✅ **Test Modu**: CATIA/Excel olmadan geliştirme ve test yapın
- ✅ **Modern GUI**: CustomTkinter ile şık ve responsive arayüz

### 🔥 Yeni Özellikler (v4.5)
- ✅ **Gelişmiş Logging**: Dosya tabanlı log sistemi (Logs/ klasörü)
- ✅ **openpyxl Desteği**: Excel okuma %300 daha hızlı!
- ✅ **Veri Doğrulama**: Parametreler CATIA'ya gönderilmeden önce validate edilir
- ✅ **Excel Şablon Oluşturma**: Tek tıkla örnek Excel dosyası oluşturun
- ✅ **HTML Rapor Export**: İşlem sonuçlarını güzel HTML raporlarına dönüştürün
- ✅ **Profil Yönetimi**: Ayarları kaydedin ve tekrar yükleyin
- ✅ **Gelişmiş Hata Yönetimi**: Detaylı hata loglama ve recovery mekanizmaları
- ✅ **Batch İşleme Optimizasyonu**: 50 satırlık batch'lerle UI dondurması yok

### 🎨 Arayüz Özellikleri
- 🎨 Modern Dark Theme
- 🐹 Hamster Wheel Animasyonu (progress indicator)
- 📊 Gerçek zamanlı istatistikler
- 🔍 Log filtreleme (Sadece hatalar/Tümü)
- ⌨️ Klavye kısayolları

## 📦 Kurulum

### Gereksinimler
- Python 3.8 veya üzeri
- Windows (CATIA entegrasyonu için)

### Adım 1: Bağımlılıkları Yükleyin

```bash
pip install -r requirements.txt
```

**Temel Gereksinimler:**
- `customtkinter` - Modern GUI framework
- `openpyxl` - Hızlı Excel okuma (ÖNERİLİR)
- `pywin32` - Windows COM (CATIA & Excel entegrasyonu için)

### Adım 2: Uygulamayı Çalıştırın

```bash
python s2dgui3.py
```

## 🎯 Kullanım Kılavuzu

### 1️⃣ Excel Dosyası Hazırlayın

Excel dosyanız şu formatta olmalı:

| ID    | Thickness | Height | P1  | D1  | P2  | D2  |
|-------|-----------|--------|-----|-----|-----|-----|
| Rib_1 | 5.0       | 20.0   | 1.5 | 0.5 | 2.0 | 0.8 |
| Rib_2 | 5.2       | 21.0   | 1.6 | 0.6 | 2.1 | 0.9 |
| Rib_3 | 5.4       | 22.0   | 1.7 | 0.7 | 2.2 | 1.0 |

**Veya** "Excel Şablon Oluştur" butonu ile otomatik şablon oluşturun!

### 2️⃣ Parametre Eşleştirmesi Yapın

1. "Ayarlar & Önizleme" sekmesine gidin
2. Her parametre için:
   - **Sütun**: Excel sütun harfi (A, B, C, ...)
   - **CATIA Parametre Adı**: ID'ye eklenecek suffix (Thickness, Height, ...)
3. Sonuç: `Rib_1Thickness`, `Rib_1Height` vb.

**Örnekler:**
- Sütun `B` → Parametre `Thickness` → CATIA'da: `Rib_1Thickness`
- Sütun `K` → Parametre `P1` → CATIA'da: `Rib_1P1`

### 3️⃣ İşlemi Başlatın

1. Excel dosyasını seçin
2. Sayfa seçin (combobox'tan)
3. "İŞLEMİ BAŞLAT ▶" butonuna tıklayın
4. Monitör sekmesinde ilerlemeyi izleyin

## ⌨️ Klavye Kısayolları

| Kısayol | Açıklama |
|---------|----------|
| `Ctrl+F` | Dosya Seç |
| `Ctrl+R` | İşlemi Başlat |
| `Ctrl+W` | İşlemi Durdur |
| `Ctrl+S` | Profil Kaydet |
| `Ctrl+O` | Profil Yükle |
| `Ctrl+E` | HTML Rapor Export |
| `Ctrl+T` | Excel Şablon Oluştur |

## 🔧 Ayarlar

### Test Modu
Dosyanın başında:
```python
TEST_MODE = True  # False yapın CATIA ile çalışmak için
```

### Profiller
- **Kaydet**: Mevcut ayarları `.json` olarak kaydedin
- **Yükle**: Daha önce kaydedilmiş ayarları yükleyin
- **Varsayılan Profil**: `default_profile.json` (varsa otomatik yüklenir)

## 📁 Klasör Yapısı

```
S2D/
├── s2dgui3.py              # Ana uygulama
├── requirements.txt        # Bağımlılıklar
├── README.md               # Bu dosya
├── default_profile.json    # Varsayılan profil (opsiyonel)
├── Logs/                   # Log dosyaları
│   └── catia_automation.log
├── Results/                # İşlem sonuç raporları (.txt)
│   └── result_YYYYMMDD_HHMMSS.txt
└── Reports/                # HTML raporları
    └── Report_YYYYMMDD_HHMMSS.html
```

## 🐛 Sorun Giderme

### "openpyxl bulunamadı" Uyarısı
```bash
pip install openpyxl
```

### "win32com bulunamadı" Hatası
```bash
pip install pywin32
```

### CATIA Bağlanamadı
- CATIA'nın açık ve bir part dosyasının aktif olduğundan emin olun
- Test modunda çalıştırarak simülasyon yapabilirsiniz

### Excel Önizlemesi Yüklenmiyor
- Dosya boyutu 100MB'ı aşmamalı
- Dosya formatı `.xlsx`, `.xlsm` veya `.xls` olmalı
- Dosyaya okuma izniniz olmalı

## 📊 Performans İyileştirmeleri

### v4.5'teki Optimizasyonlar:
- ✅ openpyxl kullanımı: %300 daha hızlı Excel okuma
- ✅ Batch UI güncelleme: 50 satırlık batch'ler (UI dondurması yok)
- ✅ Excel görünmez mod: `excel.Visible = False`
- ✅ Ekran güncellemesi kapalı: `excel.ScreenUpdating = False`
- ✅ Read-only mod: Dosyalar sadece okunur modda açılır
- ✅ Log throttling: Her 10 log'da bir render (bellek optimizasyonu)
- ✅ Max 5000 log entry (bellek sınırı)

### Performans Metrikleri:
- **50 satır**: ~2-3 saniye
- **500 satır**: ~20-30 saniye
- **5000 satır**: ~3-5 dakika

## 🔐 Güvenlik

- ✅ Read-only Excel okuma
- ✅ Dosya boyutu validasyonu (max 100MB)
- ✅ Dosya format kontrolü
- ✅ Parametre değer validasyonu
- ✅ Try-except ile güvenli hata yönetimi
- ✅ Otomatik cleanup (Excel/CATIA kapatma)

## 📝 Changelog

### v4.5 Pro (2025-11-22)
- ✨ Python logging modülü entegrasyonu
- ✨ openpyxl desteği (hızlı Excel okuma)
- ✨ Veri doğrulama sistemi
- ✨ Excel şablon oluşturma
- ✨ HTML rapor export
- ✨ Gelişmiş hata yakalama ve recovery
- ✨ Batch işleme optimizasyonu
- ✨ Klavye kısayolları genişletildi
- ✨ Profil yönetimi iyileştirildi
- 🐛 CATIA API entegrasyonu tamamlandı
- 🐛 Bellek optimizasyonları
- 🐛 UI donma sorunları düzeltildi

### v4.0
- İlk stabil sürüm
- Temel Excel → CATIA entegrasyonu
- CustomTkinter GUI
- Test modu

## 👨‍💻 Geliştirici Notları

### Kod Yapısı:
- **WorkerThread**: Arka plan işleme (threading)
- **ExcelPreviewLoader**: Önizleme yükleme (async)
- **AutomationSuite**: Ana GUI sınıfı
- **Helper Functions**: col2num, num2col, validate, vb.

### Logger Kullanımı:
```python
APP_LOGGER.info("Bilgi mesajı")
APP_LOGGER.warning("Uyarı mesajı")
APP_LOGGER.error("Hata mesajı")
APP_LOGGER.critical("Kritik hata")
```

### Yeni Özellik Ekleme:
1. Helper fonksiyonları ekleyin (başta)
2. GUI butonlarını setup_monitor/setup_settings'e ekleyin
3. Event handler metodlarını sınıfa ekleyin
4. Logger ile loglayın

## 📄 Lisans

Bu proje özel bir proje olup, ticari kullanım için izin gerektirir.

## 🤝 Destek

Sorun bildirmek veya öneride bulunmak için lütfen iletişime geçin.

---

**CATIA Automation Suite v4.5 Pro** - © 2025
*Offline, Local, Powerful* 🚀

