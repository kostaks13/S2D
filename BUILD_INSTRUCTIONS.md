# AFT Sizing Automation - EXE Build Talimatları

## Windows'ta EXE Oluşturma

### 1. Gereksinimler

```bash
pip install pyinstaller
pip install -r requirements.txt
```

### 2. Build Yöntemleri

#### Yöntem 1: Build Script Kullanma (Önerilen)

```bash
python build_exe.py
```

#### Yöntem 2: Spec Dosyası Kullanma

```bash
pyinstaller build_exe.spec
```

#### Yöntem 3: Direkt Komut

```bash
pyinstaller --name=AFT_Sizing_Automation --onefile --windowed --noconfirm --clean s2dgui4.py
```

### 3. Çıktı

EXE dosyası `dist/AFT_Sizing_Automation.exe` konumunda oluşturulacaktır.

### 4. Notlar

- **İlk çalıştırma**: Windows Defender veya antivirüs uyarısı çıkabilir (normal, çünkü imzasız)
- **Dosya boyutu**: ~50-100 MB olabilir (tüm bağımlılıklar dahil)
- **Çalışma zamanı**: İlk açılış biraz yavaş olabilir (unpacking)
- **Logs ve Results**: Bu klasörler otomatik oluşturulacak

### 5. Sorun Giderme

#### "Module not found" hatası
```bash
# Eksik modülü hidden-import olarak ekleyin
pyinstaller --hidden-import=MODULE_NAME ...
```

#### Antivirüs uyarısı
- Windows Defender'da "Daha fazla bilgi" > "Yine de çalıştır" seçin
- Veya EXE'yi imzalayın (kod imzalama sertifikası gerekir)

#### EXE çalışmıyor
- Konsol modunda çalıştırın: `--console` parametresi ekleyin
- Hata mesajlarını görmek için: `--debug=all` ekleyin

### 6. Gelişmiş Seçenekler

#### Icon ekleme
```bash
pyinstaller --icon=icon.ico s2dgui4.py
```

#### Tek klasör modu (daha hızlı başlangıç)
```bash
pyinstaller --onedir --windowed s2dgui4.py
```

#### UPX ile sıkıştırma (daha küçük dosya)
```bash
pyinstaller --upx-dir=/path/to/upx --onefile s2dgui4.py
```

