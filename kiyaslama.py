import streamlit as st
import pandas as pd
import re
import io

# Sayfa ayarları
st.set_page_config(page_title="Fiyat Karşılaştırma Sistemi", layout="wide")

st.title("📊 Profesyonel Fiyat Analiz ve Raporlama")
st.markdown("Güncel kurları giriniz ve hücre bazlı döviz içeren Excel dosyasını yükleyiniz.")

# Kur girişleri
col1, col2 = st.columns(2)
usd_kur = col1.number_input("Güncel USD Kuru (TL)", value=44.00, step=0.1)
eur_kur = col2.number_input("Güncel EUR Kuru (TL)", value=52.00, step=0.1)

uploaded_file = st.file_uploader("Analiz Edilecek Excel Dosyasını Yükleyiniz", type=["xlsx", "xls"])

# Hücreden hem rakamı hem de kuru ayrıştıran fonksiyon
def kur_ve_fiyat_ayristir(deger):
    if pd.isna(deger): return None, None
    
    deger_str = str(deger).upper().replace(',', '.')
    
    # Kuru tespit et
    kur_tipi = "TL" # Varsayılan
    if any(x in deger_str for x in ["$", "USD"]):
        kur_tipi = "USD"
    elif any(x in deger_str for x in ["€", "EUR"]):
        kur_tipi = "EUR"
    elif any(x in deger_str for x in ["₺", "TL"]):
        kur_tipi = "TL"
        
    # Sadece rakamı çek
    rakam_match = re.search(r"(\d+\.?\d*)", deger_str)
    if rakam_match:
        return float(rakam_match.group(1)), kur_tipi
    return None, None

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # İlk 6 sütun (Sıra, Kod, Ad, Açıklama, Miktar, Birim) sabit kabul edilir.
    # Sonraki tüm sütunlar tedarikçi olarak değerlendirilir.
    tedarikci_sutunlari = df.columns[6:].tolist()
    
    if not tedarikci_sutunlari:
        st.warning("Hata: Tedarikçi verisi içeren sütun bulunamadı.")
    else:
        sonuc_df = df.copy()
        en_ucuz_tedarikciler = []
        en_ucuz_fiyatlar = []
        
        for index, row in sonuc_df.iterrows():
            min_tl = float('inf')
            en_iyi_firma, en_iyi_fiyat_metni = "-", "-"
            
            for col in tedarikci_sutunlari:
                sayi, birim = kur_ve_fiyat_ayristir(row[col])
                
                if sayi is not None:
                    # TL karşılığını hesapla
                    kur_degeri = usd_kur if birim == "USD" else (eur_kur if birim == "EUR" else 1)
                    tl_karsiligi = sayi * kur_degeri
                    
                    if tl_karsiligi < min_tl:
                        min_tl = tl_karsiligi
                        en_iyi_firma = col
                        en_iyi_fiyat_metni = f"{sayi} {birim}"
            
            en_ucuz_tedarikciler.append(en_iyi_firma)
            en_ucuz_fiyatlar.append(en_iyi_fiyat_metni)
            
        sonuc_df['En Uygun Tedarikçi'] = en_ucuz_tedarikciler
        sonuc_df['En Uygun Fiyat'] = en_ucuz_fiyatlar
        
        st.write("### Analiz Önizleme")
        st.dataframe(sonuc_df)

        # --- Excel Tasarımı ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sonuc_df.to_excel(writer, index=False, sheet_name='Analiz Raporu')
            workbook = writer.book
            worksheet = writer.sheets['Analiz Raporu']

            header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'fg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
            result_fmt = workbook.add_format({'bold': True, 'fg_color': '#FFEB9C', 'border': 1})
            border_fmt = workbook.add_format({'border': 1})

            last_col = len(sonuc_df.columns) - 1
            row_count = len(sonuc_df)

            for col_num, value in enumerate(sonuc_df.columns.values):
                worksheet.write(0, col_num, value, header_fmt)
                worksheet.set_column(col_num, col_num, 18)

            # Sadece dolu satırları formatla
            worksheet.conditional_format(1, 0, row_count, last_col - 2, {'type': 'no_errors', 'format': border_fmt})
            worksheet.conditional_format(1, last_col - 1, row_count, last_col, {'type': 'no_errors', 'format': result_fmt})
            worksheet.freeze_panes(1, 0)

        st.download_button(
            label="📥 Raporu Excel Olarak İndir",
            data=output.getvalue(),
            file_name="Fiyat_Karsilastirma_Raporu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )