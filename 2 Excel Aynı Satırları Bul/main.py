"""
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
"""

import os
import glob
import pandas as pd



# =======================================================================
#  AYARLAR  —  Burası düzenlenir
# =======================================================================

DIZIN = os.path.dirname(os.path.abspath(__file__))  # script ile aynı klasör

A_ISIM = "a"   # A dosyasının adı (uzantısız)
B_ISIM = "b"   # B dosyasının adı (uzantısız)

# A dosyasından arama yapılacak sütun:
#   Harf ile     →  A_SUTUN = "E"        (büyük/küçük fark etmez)
#   İsim ile     →  A_SUTUN = "Müşteri"
#   İndeks ile   →  A_SUTUN = 4          (0 tabanlı; E = 4)
A_SUTUN = "V"

# B dosyasında hangi sütun(lar)da aranacak:
#   Tek sütun    →  B_SUTUNLAR = "SiteNo"
#   Çok sütun    →  B_SUTUNLAR = ["SiteNo", "Adres", "ID"]
#   Tumu         →  B_SUTUNLAR = None   ← bütün sütunlarda ara
B_SUTUNLAR = None

# --- CSV seçenekleri (otomatik tespitte başarısız olunursa kullanılır) ---
A_CSV_ENCODING  = None   # Örnek: "cp1254" | None → otomatik dene
A_CSV_SEPARATOR = None   # Örnek: ";"       | None → otomatik dene

B_CSV_ENCODING  = None
B_CSV_SEPARATOR = None

# Çıktı dosyası adları (uzantı kaynak dosyayla aynı olur)
CIKTI_VAR  = "a_var"   # A'da var, B'de bulunan satırlar
CIKTI_YOK  = "a_yok"  # A'da var, B'de bulunmayan satırlar

# =======================================================================


# -----------------------------------------------------------------------
#  Yardımcı fonksiyonlar
# -----------------------------------------------------------------------

def dosya_bul(dizin: str, isim: str) -> str:
    """Dizinde 'isim' adlı xls / xlsx / csv dosyasını bulur."""
    for uzanti in ("xlsx", "xls", "csv"):
        eslesme = glob.glob(os.path.join(dizin, f"{isim}.{uzanti}"))
        if eslesme:
            return eslesme[0]
    raise FileNotFoundError(
        f"'{isim}' adında xls/xlsx/csv dosyası bulunamadı: {dizin}"
    )


def _csv_dene(yol: str, encoding: str | None, separator: str | None) -> pd.DataFrame:
    """
    Encoding ve/veya separator verilmişse doğrudan dener;
    None ise birden fazla kombinasyonu sırayla dener.
    """
    enc_listesi   = [encoding]   if encoding  else ["cp1254", "utf-8-sig", "WINDOWS-1252", "latin-1"]
    sep_listesi   = [separator]  if separator else [";", ",", "\t", "|"]

    for enc in enc_listesi:
        for sep in sep_listesi:
            try:
                df = pd.read_csv(yol, dtype=str, encoding=enc, sep=sep)
                # Tek sütunlu gelirse yanlış ayırıcı denemesi olabilir
                if df.shape[1] > 1:
                    return df
            except Exception:
                continue

    # Son çare: tek sütunlu da olsa oku
    for enc in enc_listesi:
        for sep in sep_listesi:
            try:
                return pd.read_csv(yol, dtype=str, encoding=enc, sep=sep)
            except Exception:
                continue

    raise ValueError(f"CSV dosyası okunamadı: {yol}")


def dosya_oku(yol: str, encoding: str | None = None, separator: str | None = None) -> pd.DataFrame:
    """Dosya uzantısına göre DataFrame olarak okur."""
    uzanti = os.path.splitext(yol)[1].lower()
    if uzanti == ".csv":
        df = _csv_dene(yol, encoding, separator)
        print(f"   CSV okundu  →  sütun sayısı: {df.shape[1]}, satır: {df.shape[0]}")
        return df
    elif uzanti == ".xls":
        return pd.read_excel(yol, dtype=str, engine="xlrd")
    else:
        return pd.read_excel(yol, dtype=str, engine="openpyxl")


def sutun_sec(df: pd.DataFrame, secim, dosya_etiketi: str) -> str:
    """
    secim:
      - str  → sütun harfi (A-Z) veya sütun başlığı
      - int  → 0 tabanlı sütun indeksi
    Gerçek sütun adını (df.columns'daki) döner.
    """
    kolonlar = list(df.columns)

    if isinstance(secim, int):
        if secim < 0 or secim >= len(kolonlar):
            raise IndexError(
                f"{dosya_etiketi}: indeks {secim} geçersiz "
                f"(toplam {len(kolonlar)} sütun var)."
            )
        return kolonlar[secim]

    # String: önce sütun harfi mi kontrol et (A, B, ... Z, AA, ...)
    harf = secim.strip().upper()
    harf_indeks = harf_to_index(harf)
    if harf_indeks is not None:
        if harf_indeks >= len(kolonlar):
            raise IndexError(
                f"{dosya_etiketi}: '{harf}' sütunu yok "
                f"(toplam {len(kolonlar)} sütun var)."
            )
        return kolonlar[harf_indeks]

    # Sonra sütun başlığı olarak ara
    if secim in kolonlar:
        return secim

    # Büyük/küçük harf duyarsız arama
    alt_secim  = secim.strip().lower()
    alt_kolonlar = [k.strip().lower() for k in kolonlar]
    if alt_secim in alt_kolonlar:
        return kolonlar[alt_kolonlar.index(alt_secim)]

    raise KeyError(
        f"{dosya_etiketi}: '{secim}' sütunu bulunamadı.\n"
        f"Mevcut sütunlar: {kolonlar}"
    )


def harf_to_index(harf: str) -> int | None:
    """Excel sütun harfini 0 tabanlı indekse çevirir. Harf değilse None döner."""
    harf = harf.upper()
    if not all(c.isalpha() for c in harf):
        return None
    sonuc = 0
    for c in harf:
        sonuc = sonuc * 26 + (ord(c) - ord('A') + 1)
    return sonuc - 1


def deger_temizle(v) -> str:
    """
    Değeri karşılaştırma için normalize eder:
    - NaN → boş string
    - Baştaki/sondaki boşlukları siler
    - XLS'in metin-sayı geçici apostrof'unu ('141000...) kaldırır
    """
    if not isinstance(v, str) and pd.isna(v):
        return ""
    return str(v).strip().strip("'\"")


def kaydet(df: pd.DataFrame, kaynak_yol: str, cikti_isim: str) -> str:
    """Sonucu kaynak dosyayla aynı dizine ve aynı uzantıyla kaydeder."""
    dizin   = os.path.dirname(kaynak_yol)
    uzanti  = os.path.splitext(kaynak_yol)[1].lower()

    # .xls yazmak modern pandas'ta desteklenmiyor → xlsx'e yükselt
    if uzanti == ".xls":
        uzanti = ".xlsx"

    cikti   = os.path.join(dizin, f"{cikti_isim}{uzanti}")

    if uzanti == ".csv":
        df.to_csv(cikti, index=False, encoding="utf-8-sig", sep=";")
    else:
        df.to_excel(cikti, index=False)

    return cikti


# -----------------------------------------------------------------------
#  Ana akış
# -----------------------------------------------------------------------

def main():
    a_yol = dosya_bul(DIZIN, A_ISIM)
    b_yol = dosya_bul(DIZIN, B_ISIM)
    print(f"\nA dosyası : {a_yol}")
    print(f"B dosyası : {b_yol}\n")

    df_a = dosya_oku(a_yol, A_CSV_ENCODING, A_CSV_SEPARATOR)
    df_b = dosya_oku(b_yol, B_CSV_ENCODING, B_CSV_SEPARATOR)

    # A'da arama yapılacak sütunu belirle
    a_sutun_adi = sutun_sec(df_a, A_SUTUN, "A dosyası")
    print(f"A dosyasında aranan sütun  : '{a_sutun_adi}'")

    # B'de hangi sütunlara bakılacak
    if B_SUTUNLAR is None:
        b_sutun_listesi = list(df_b.columns)
        print(f"B dosyasında arama         : TÜM SÜTUNLAR ({len(b_sutun_listesi)} adet)")
    elif isinstance(B_SUTUNLAR, (str, int)):
        b_sutun_listesi = [sutun_sec(df_b, B_SUTUNLAR, "B dosyası")]
        print(f"B dosyasında arama         : '{b_sutun_listesi[0]}'")
    else:
        b_sutun_listesi = [sutun_sec(df_b, s, "B dosyası") for s in B_SUTUNLAR]
        print(f"B dosyasında arama         : {b_sutun_listesi}")

    # B'deki hedef değerleri küme olarak topla (O(1) arama)
    b_degerler: set = set()
    for sutun in b_sutun_listesi:
        b_degerler.update(
            df_b[sutun].dropna().apply(deger_temizle)
        )

    print(f"B'de toplam benzersiz değer: {len(b_degerler)}\n")

    # Eşleşme kontrolü
    a_var_satirlar = []   # B'de bulunan A satırları
    a_yok_satirlar = []   # B'de bulunmayan A satırları

    for _, satir in df_a.iterrows():
        deger = deger_temizle(satir[a_sutun_adi])
        if deger and deger in b_degerler:
            a_var_satirlar.append(satir)
        else:
            a_yok_satirlar.append(satir)

    df_a_var = pd.DataFrame(a_var_satirlar, columns=df_a.columns)  # A'nın B'de bulunan satırları
    df_a_yok = pd.DataFrame(a_yok_satirlar, columns=df_a.columns)  # A'nın B'de bulunmayan satırları

    # Kaydet
    print("-" * 55)
    if not df_a_var.empty:
        yol = kaydet(df_a_var, a_yol, CIKTI_VAR)
        print(f"✔  B'de bulunan    {len(df_a_var):>5} satır (A'dan)  →  {yol}")
    else:
        print("ℹ  B'de eşleşen A kaydı yok; 'a_var' dosyası oluşturulmadı.")

    if not df_a_yok.empty:
        yol = kaydet(df_a_yok, a_yol, CIKTI_YOK)
        print(f"✔  B'de bulunmayan {len(df_a_yok):>5} satır (A'dan)  →  {yol}")
    else:
        print("ℹ  Tüm A kayıtları B'de eşleşti; 'a_yok' dosyası oluşturulmadı.")
    print("-" * 55)


if __name__ == "__main__":
    main()
