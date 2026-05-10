# 📧 Mail Analyzer Pro

**Mail Analyzer Pro**, Microsoft Outlook posta kutunuzdaki ve bağlı PST/arşiv dosyalarınızdaki tüm klasör yapısını görselleştirmenizi, mail içeriklerinde anahtar kelime araması yapmanızı ve arama sonuçlarına ait istatistikleri (kaç mail, kimden, hangi dönemde) görmenizi sağlayan yerel bir masaüstü uygulamasıdır.

> **Gereksinim:** Windows + Microsoft Outlook (klasik masaüstü sürümü) + Python 3.10+

---

## 🚀 Kurulum

```bash
# 1. Depoyu klonlayın veya dosyaları indirin
# 2. Sanal ortam oluşturun (önerilir)
python -m venv venv
venv\Scripts\activate

# 3. Bağımlılıkları yükleyin
pip install -r requirements.txt

# 4. Uygulamayı başlatın
python gui_app.py
```

---

## 🖥️ Ekranlar ve Nasıl Çalışır

### 📁 Explorer — Klasör Ağacı

Uygulama açıldığında sol panelde Outlook'taki **tüm posta kutuları ve PST dosyalarının iç içe klasör yapısı** ağaç (tree) şeklinde görünür.

- Her store (posta kutusu / PST) en üst düğüm olarak listelenir.
- Altında klasörler ve alt klasörler hiyerarşik olarak sıralanır.
- Her klasörün yanında o klasördeki **mail sayısı** gösterilir.
- Bir klasöre tıkladığınızda sağ panelde o klasördeki **mailler anında listelenir** (konu, gönderen, tarih).
- Bir maile **çift tıkladığınızda** tam içerik (konu, gönderen, tarih, gövde metni) ayrı bir pencerede açılır.
- **"Seçili Klasörü Tara & Kaydet"** butonu ile seçili klasördeki mailler yerel veritabanına kaydedilir ve risk analizi yapılır.

```
┌─────────────────────────┐   ┌──────────────────────────────────────┐
│ 🗄 Şirket Postası       │   │  Konu          │ Gönderen  │ Tarih   │
│  ├─ Gelen Kutusu  (312) │   │  Toplantı Agn. │ ali@...   │ 2024-01 │
│  ├─ Gönderilen    (87)  │──▶│  Proje Raporu  │ ayse@...  │ 2024-01 │
│  └─ Arşiv               │   │  ...           │           │         │
│ 🗄 Arşiv.pst            │   └──────────────────────────────────────┘
│  └─ 2023 Mailleri (1.2k)│
└─────────────────────────┘
```

---

### 🔍 Arama — Canlı Mail Arama

**Arama** ekranında Outlook'taki tüm posta kutuları ve PST'ler **canlı olarak** aranır; veritabanı taraması değil, doğrudan Outlook üzerinden yapılır.

| Alan | Açıklama |
|---|---|
| **Konu içerir** | Mail konusunda geçen metin |
| **Gönderen** | Gönderenin e-posta adresi (kısmi eşleşme) |
| **İçerik (anahtar kelime)** | Mail gövdesinde aranacak kelime |

Alanları doldurup **ARA** butonuna basın:

1. Arama tüm store ve klasörlerde eş zamanlı olarak çalışır.
2. Eşleşen mailler alt tabloda listelenir.
3. Tablonun üstünde **toplam kaç sonuç** ve **kaç farklı gönderen** olduğu gösterilir.
4. Sonuçlardan birine **çift tıklayarak** tam mail içeriğini okuyabilirsiniz.

---

### 📊 İstatistik — Sayısal Özet

**İstatistik** ekranı, daha önce **Tara & Kaydet** ile veritabanına aktarılmış mailler üzerinden anlık sayısal özet sunar.

| Kart | Gösterge |
|---|---|
| Toplam Mail | DB'ye kaydedilmiş toplam mail sayısı |
| Yüksek Risk | Risk skoru ≥ 50 olan mail sayısı |
| Store Sayısı | Kaç farklı posta kutusu / PST tarandı |
| Klasör Sayısı | Kaç farklı klasör tarandı |

Aşağısında **"En Çok Gönderen"** tablosu yer alır — hangi adresten kaç mail geldiğini gösterir.

---

### 📋 Sonuçlar — Risk Logu & Rapor

**Sonuçlar** ekranında risk skoruna göre sıralanmış mail listesi görünür.

- **Risk skoru**, mail içindeki şüpheli anahtar kelimeler (password, vpn, admin, token…) ve Regex eşleşmelerine (IP adresi, JWT token, e-posta adresi) göre otomatik hesaplanır.
- **📥 XLSX Rapor Al** butonu ile tüm analiz çıktısı `reports/` klasörüne Excel olarak kaydedilir.

---

## 🗂️ Proje Yapısı

```
mail-analyzer/
├── gui_app.py                  # Ana uygulama (GUI)
├── requirements.txt
├── assets/
│   ├── icon.ico
│   └── icon.png
├── database/
│   └── mail_analyzer.db        # Yerel SQLite veritabanı (otomatik oluşur)
├── reports/                    # XLSX raporlar (otomatik oluşur)
└── backend/app/
    ├── core/database.py        # SQLite bağlantı & tablo yönetimi
    ├── models/mail.py          # Mail veri modeli
    ├── services/
    │   ├── outlook_service.py  # Outlook COM köprüsü (klasör/arama/çekme)
    │   ├── analyzer_service.py # Risk skoru hesaplama
    │   ├── keyword_service.py  # Anahtar kelime eşleştirme
    │   ├── regex_service.py    # Regex tabanlı tespit
    │   └── report_service.py  # Excel raporlama
    └── workers/
        └── scan_worker.py      # Mail → DB kayıt işçisi
```

---

## ⚙️ Teknik Notlar

- **COM Güvenliği:** Outlook ile her arka plan thread'i kendi `CoInitialize / CoUninitialize` döngüsünü yönetir. GUI donmaz.
- **Veritabanı:** SQLite (`database/mail_analyzer.db`). Harici kurulum gerektirmez.
- **Aynı maili tekrar tarama:** `message_id` (EntryID) üzerinden `MERGE` yapılır; çift kayıt oluşmaz.
- **Gövde boyutu sınırı:** Performans için her mailin gövdesi 20.000 karakterle sınırlandırılmıştır.

---

## 🛡️ Lisans

Bu proje güvenlik analizi ve adli bilişim (forensic) amaçları için geliştirilmiştir.
