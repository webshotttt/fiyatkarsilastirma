import streamlit as st
import pandas as pd
import openpyxl
import re
import io

# Sayfa ayarları
st.set_page_config(page_title="Fiyat Karşılaştırma Sistemi", layout="wide")

st.title("📊 Profesyonel Fiyat Analiz ve Raporlama")
st.markdown("Güncel kurları giriniz ve analiz edilecek veri setini yükleyiniz.")

# Kur girişleri
col1, col2 = st.columns(2)
usd_kur = col1.number_input("Güncel USD Kuru (TL)", value=44.00, step=0.1)
eur_kur = col2.number_input("Güncel EUR Kuru (TL)", value=52.00, step=0.1)

uploaded_file = st.file_uploader("Analiz Edilecek Excel Dosyasını Yükleyiniz", type=["xlsx", "xls"])

if uploaded_file:
    # 1. Pandas ile tablonun iskeletini oku
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)
    
    # 2. Openpyxl ile arka plandaki hücre formatlarını (₺, $, €) okumak için yükle
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active # Aktif sayfayı al
    
    # İlk 6 sütun (Sıra, Kod, Ad, Açıklama, Miktar, Birim) sabit. Sonrakiler tedarikçi.
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
            
            # Pandas satır indeksi 0'dan başlar, Excel'de veriler 2. satırdan (1. satır başlıktır)
            excel_row = index + 2 
            
            for col_name in tedarikci_sutunlari:
                # Sütunun Excel'deki harf/sayı indeksini bul
                excel_col = df.columns.get_loc(col_name) + 1
                
                # İlgili hücreyi openpyxl ile nokta atışı yakala
                cell = ws.cell(row=excel_row, column=excel_col)
                sayi = cell.value
                
                if sayi is not None and str(sayi).strip() != "":
                    # 1. Sayıyı Rakam Olarak Çekme
                    if isinstance(sayi, (int, float)):
                        sayi_float = float(sayi)
                    else:
                        # Eğer hücre metin ise ("1,50 USD" gibi yazılmışsa)
                        deger_str = str(sayi).replace(',', '.')
                        rakam_match = re.search(r"(\d+\.?\d*)", deger_str)
                        if rakam_match:
                            sayi_float = float(rakam_match.group(1))
                        else:
                            sayi_float = None
                            
                    if sayi_float is not None:
                        # 2. Para Birimini Çekme (Hem Formata Hem Metne Bakar)
                        format_str = str(cell.number_format).upper()
                        hucre_metni = str(sayi).upper()
                        
                        birim = "TL" # Varsayılan kabul
                        
                        if any(x in format_str for x in ["$", "USD"]) or any(x in hucre_metni for x in ["$", "USD"]):
                            birim = "USD"
                        elif any(x in format_str for x in ["€", "EUR"]) or any(x in hucre_metni for x in ["€", "EUR"]):
                            birim = "EUR"
                        elif any(x in format_str for x in ["₺", "TL", "TRY"]) or any(x in hucre_metni for x in ["₺", "TL", "TRY"]):
                            birim = "TL"
                            
                        # 3. Kuru TL'ye Çevirip Karşılaştırma
                        kur_degeri = usd_kur if birim == "USD" else (eur_kur if birim == "EUR" else 1)
                        tl_karsiligi = sayi_float * kur_degeri
                        
                        if tl_karsiligi < min_tl:
                            min_tl = tl_karsiligi
                            en_iyi_firma = col_name
                            # Finansal formatta okuyunca 1.0 yerine 1 çıkmaması için
                            en_iyi_fiyat_metni = f"{sayi_float} {birim}"
            
            en_ucuz_tedarikciler.append(en_iyi_firma)
            en_ucuz_fiyatlar.append(en_iyi_fiyat_metni)
            
        sonuc_df['En Uygun Tedarikçi'] = en_ucuz_tedarikciler
        sonuc_df['En Uygun Fiyat'] = en_ucuz_fiyatlar
        
        st.write("### Analiz Önizleme")
        st.dataframe(sonuc_df)

        # --- Excel Tasarımı Bölümü ---
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
                # Sütun genişliklerini ayarla
                if col_num >= last_col - 1:
                    worksheet.set_column(col_num, col_num, 20)
                else:
                    worksheet.set_column(col_num, col_num, max(len(str(value)), 15))

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