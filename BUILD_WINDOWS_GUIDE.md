# macOS'ta Windows EXE Oluşturma Rehberi

macOS'ta Windows için EXE oluşturmanın birkaç yolu var:

## 🚀 Yöntem 1: GitHub Actions (ÖNERİLEN - En Kolay)

### Avantajlar:
- ✅ Ücretsiz
- ✅ Otomatik
- ✅ Windows'ta gerçek build
- ✅ Her push'ta otomatik build

### Adımlar:

1. **GitHub'a projeyi yükleyin:**
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/KULLANICI_ADI/REPO_ADI.git
git push -u origin main
```

2. **GitHub'da Actions sekmesine gidin**

3. **"Build Windows EXE" workflow'unu çalıştırın**

4. **EXE dosyasını indirin** (Artifacts'tan)

---

## 💻 Yöntem 2: Windows Virtual Machine

### Gereksinimler:
- Parallels Desktop, VMware Fusion veya VirtualBox
- Windows 10/11 ISO

### Adımlar:

1. **Windows VM kurun**

2. **Projeyi VM'e kopyalayın:**
   - Shared folder kullanın
   - Veya git clone yapın

3. **VM'de build yapın:**
```bash
# VM'de PowerShell veya CMD açın
pip install -r requirements.txt
pip install pyinstaller
python build_exe.py
```

---

## ☁️ Yöntem 3: Windows Bulut Servisi

### Seçenekler:
- **AWS EC2** (Windows Server)
- **Azure Virtual Machines**
- **Google Cloud Compute Engine**

### Avantajlar:
- Gerçek Windows ortamı
- İstediğiniz zaman kullan

### Dezavantajlar:
- Ücretli (saatlik)
- Kurulum gerekir

---

## 🐳 Yöntem 4: Docker (Gelişmiş)

Windows container kullanarak (daha karmaşık):

```dockerfile
FROM mcr.microsoft.com/windows/servercore:ltsc2022
# Python ve PyInstaller kurulumu
```

---

## 📋 Hızlı Başlangıç - GitHub Actions

1. `.github/workflows/build-windows.yml` dosyası zaten hazır
2. GitHub'a push yapın
3. Actions sekmesinden "Run workflow" tıklayın
4. EXE'yi indirin

---

## ⚠️ Önemli Notlar

- **Cross-compilation çalışmaz**: macOS'ta direkt Windows EXE oluşturamazsınız
- **En pratik çözüm**: GitHub Actions (ücretsiz ve otomatik)
- **Test için**: Windows VM kullanın

---

## 🔧 Manuel Build (Windows'ta)

Windows'ta olduğunuzda:

```bash
# 1. Gereksinimleri yükle
pip install -r requirements.txt
pip install pyinstaller

# 2. Build yap
python build_exe.py

# VEYA
pyinstaller build_exe.spec
```

EXE dosyası `dist/AFT_Sizing_Automation.exe` konumunda olacak.

