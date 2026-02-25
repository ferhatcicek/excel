==============================================================================
  2 Excel Aynı Satırları Bul  —  Kullanım Kılavuzu
==============================================================================

AMACI:
  "a" adlı dosyanın belirtilen sütunundaki değerleri "b" adlı dosyada arar.
    • Eşleşen  A satırları  →  a_var  (örn. a_var.xlsx)
    • Eşleşmeyen A satırları  →  a_yok  (örn. a_yok.xlsx)

DESTEKLENEN DOSYA FORMATLARI:
  .xlsx  .xls  .csv  (a ve b dosyaları farklı formatlarda olabilir)

KURULUM (ilk kullanımda bir kere yapılır):
  pip install pandas openpyxl xlrd

ÇALIŞTIRMA:
  python main.py

AYARLAR (bu dosyanın "AYARLAR" bölümünü düzenleyin):
  ┌──────────────────┬──────────────────────────────────────────────────────┐
  │ Değişken         │ Açıklama & Örnekler                                  │
  ├──────────────────┼──────────────────────────────────────────────────────┤
  │ DIZIN            │ Dosyaların bulunduğu klasör.                         │
  │                  │ Varsayılan: script ile aynı klasör.                  │
  │                  │ Örnek: r"D:\\veri\\klasor"                           │
  ├──────────────────┼──────────────────────────────────────────────────────┤
  │ A_ISIM / B_ISIM  │ Dosya adları (uzantısız).                            │
  │                  │ Örnek: A_ISIM = "musteri_listesi"                    │
  ├──────────────────┼──────────────────────────────────────────────────────┤
  │ A_SUTUN          │ A dosyasında arama yapılacak sütun:                  │
  │                  │   Harf   → "E"  (Excel sütun harfi)                  │
  │                  │   İsim   → "Hizmet Numarası"                         │
  │                  │   İndeks → 4   (0 tabanlı; A=0, B=1, ... E=4)       │
  ├──────────────────┼──────────────────────────────────────────────────────┤
  │ B_SUTUNLAR       │ B dosyasında aramanın yapılacağı sütun(lar):         │
  │                  │   None            → tüm sütunlarda ara               │
  │                  │   "SiteNo"        → sadece bu sütunda ara            │
  │                  │   ["A", "SiteNo"] → bu sütunlarda ara                │
  ├──────────────────┼──────────────────────────────────────────────────────┤
  │ A_CSV_ENCODING   │ A CSV ise encoding (None = otomatik dene)            │
  │ A_CSV_SEPARATOR  │ A CSV ise ayırıcı  (None = otomatik dene)            │
  │ B_CSV_ENCODING   │ B CSV ise encoding  Örnek: "cp1254"                  │
  │ B_CSV_SEPARATOR  │ B CSV ise ayırıcı   Örnek: ";"  ","  "\\t"           │
  ├──────────────────┼──────────────────────────────────────────────────────┤
  │ CIKTI_VAR        │ Eşleşen satırların yazılacağı dosya adı (uzantısız)  │
  │ CIKTI_YOK        │ Eşleşmeyen satırların dosya adı (uzantısız)          │
  └──────────────────┴──────────────────────────────────────────────────────┘

NOT: Büyük sayı içeren hücreler XLS'te otomatik apostrof (') ile saklanır.
     Bu script karşılaştırma öncesinde apstrofu otomatik temizler.
==============================================================================
