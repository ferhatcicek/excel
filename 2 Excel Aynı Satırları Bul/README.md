# 2 Excel — Aynı Satırları Bul

İki Excel/CSV dosyasını karşılaştırır: **A** dosyasının seçili sütunundaki değerleri **B** dosyasında arar ve sonuçları ayrı dosyalara yazar.

```
A dosyası (seçili sütun)  ──►  B dosyasında ara
                                      │
               ┌──────────────────────┴──────────────────────┐
           BULUNDU                                       BULUNAMADI
               │                                             │
          a_var.xlsx                                    a_yok.xlsx
      (A'nın eşleşen satırları)               (A'nın eşleşmeyen satırları)
```

---

## Özellikler

- `.xlsx`, `.xls`, `.csv` formatlarını destekler; A ve B farklı formatlarda olabilir
- Sütun seçimi: **Excel harfi** (`"E"`), **başlık adı** (`"Hizmet No"`) veya **0 tabanlı indeks** (`4`)
- B dosyasında **tek sütun**, **birden fazla sütun** veya **tüm sütunlarda** arama
- CSV için encoding ve separator **otomatik tespit** (manuel atama da desteklenir)
- XLS'teki apostrof (`'`) ile saklanan büyük sayılar otomatik normalize edilir

---

## Kurulum

```bash
pip install pandas openpyxl xlrd
```

---

## Kullanım

1. `a` ve `b` adlı dosyalarınızı script ile **aynı klasöre** koyun.
2. `main.py` dosyasının üst kısmındaki **AYARLAR** bölümünü düzenleyin.
3. Çalıştırın:

```bash
python main.py
```

Çıktı dosyaları (`a_var.xlsx` / `a_yok.xlsx`) aynı klasörde oluşur.

---

## Ayarlar

`main.py` dosyasının en üstündeki `# AYARLAR` bölümü:

```python
DIZIN = os.path.dirname(os.path.abspath(__file__))  # Varsayılan: script klasörü

A_ISIM = "a"   # A dosyasının adı (uzantısız)
B_ISIM = "b"   # B dosyasının adı (uzantısız)

# A'da aranan sütun — üç yöntemden biri:
A_SUTUN = "E"            # Excel harfi
# A_SUTUN = "Hizmet No"  # Sütun başlığı
# A_SUTUN = 4            # 0 tabanlı indeks

# B'de arama yapılacak sütun(lar):
B_SUTUNLAR = None          # None → tüm sütunlar
# B_SUTUNLAR = "SiteNo"   # Tek sütun
# B_SUTUNLAR = ["A", "B"] # Çok sütun

# CSV encoding / separator (None = otomatik):
A_CSV_ENCODING  = None
A_CSV_SEPARATOR = None
B_CSV_ENCODING  = None
B_CSV_SEPARATOR = None

CIKTI_VAR = "a_var"   # Eşleşen satırların dosya adı
CIKTI_YOK = "a_yok"   # Eşleşmeyen satırların dosya adı
```

---

## Gereksinimler

| Paket | Amaç |
|---|---|
| `pandas` | Veri okuma/yazma |
| `openpyxl` | `.xlsx` okuma/yazma |
| `xlrd` | `.xls` okuma |
