import streamlit as st
import pandas as pd
import openpyxl
import re
import io
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime

# Sayfa ayarları
st.set_page_config(page_title="Fiyat Karşılaştırma Sistemi", layout="wide")

st.title("📊 Profesyonel Fiyat Analiz ve Raporlama")
st.markdown("Kurlar **TCMB (Merkez Bankası)** üzerinden canlı çekilmektedir. Analiz edilecek veri setini yükleyiniz.")

# Arka planda 60 saniyede bir güncellenen Merkez Bankası Kur Motoru
@st.cache_data(ttl=60)
def canli_kurlari_cek():
    try:
        # TCMB Canlı XML adresi
        url = "https://www.tcmb.gov.tr/kurlar/today.xml"
        # Bot gibi görünmemek için tarayıcı kılığına giriyoruz
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            xml_data = response.read()

        root = ET.fromstring(xml_data)

        # TCMB Döviz Satış (ForexSelling) kurlarını çekiyoruz
        usd_kur = float(root.find(".//Currency[@CurrencyCode='USD']/ForexSelling").text)
        eur_kur = float(root.find(".//Currency[@CurrencyCode='EUR']/ForexSelling").text)

        guncelleme_zamani = datetime.now().strftime("%d.%m.%Y - %H:%M")
        return round(usd_kur, 4), round(eur_kur, 4), guncelleme_zamani
    except Exception as e:
        return None, None, None

# Kurları çek ve ekrana yazdır
canli_usd, canli_eur, son_guncelleme = canli_kurlari_cek()

if canli_usd and canli_eur:
    st.success(f"✅ Kurlar TCMB'den Canlı Çekildi. (Güncellenme Tarihi: {son_guncelleme})")
    default_usd = float(canli_usd)
    default_eur = float(canli_eur)
else:
    st.warning("⚠️ Canlı kur sunucusuna ulaşılamadı! Lütfen kurları manuel giriniz.")
    default_usd = 44.00
    default_eur = 52.00

# Kur girişleri
col1, col2 = st.columns(2)
usd_kur = col1.number_input("Güncel USD Kuru (TL)", value=default_usd, step=0.01, format="%.4f")
eur_kur = col2.number_input("Güncel EUR Kuru (TL)", value=default_eur, step=0.01, format="%.4f")

st.markdown("---")

uploaded_file = st.file_uploader("Analiz Edilecek Excel Dosyasını Yükleyiniz", type=["xlsx", "xls"])

if uploaded_file:
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)
    
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active 
    
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
            
            excel_row = index + 2 
            
            for col_name in tedarikci_sutunlari:
                excel_col = df.columns.get_loc(col_name) + 1
                cell = ws.cell(row=excel_row, column=excel_col)
                sayi = cell.value
                
                if sayi is not None and str(sayi).strip() != "":
                    if isinstance(sayi, (int, float)):
                        sayi_float = float(sayi)
                    else:
                        deger_str = str(sayi).replace(',', '.')
                        rakam_match = re.search(r"(\d+(?:\.\d+)?)", deger_str)
                        if rakam_match:
                            sayi_float = float(rakam_match.group(1))
                        else:
                            sayi_float = None
                            
                    if sayi_float is not None:
                        format_str = str(cell.number_format).upper()
                        hucre_metni = str(sayi).upper()
                        
                        kombine_metin = f"{format_str} {hucre_metni}"
                        birim = "TL"
                        
                        if any(x in kombine_metin for x in ["€", "EUR", "EURO"]):
                            birim = "EUR"
                        elif any(x in kombine_metin for x in ["₺", "TL", "TRY"]):
                            birim = "TL"
                        elif any(x in kombine_metin for x in ["$", "USD"]):
                            birim = "USD"
                            
                        kur_degeri = usd_kur if birim == "USD" else (eur_kur if birim == "EUR" else 1)
                        tl_karsiligi = sayi_float * kur_degeri
                        
                        if tl_karsiligi < min_tl:
                            min_tl = tl_karsiligi
                            en_iyi_firma = col_name
                            en_iyi_fiyat_metni = f"{sayi_float} {birim}"
            
            en_ucuz_tedarikciler.append(en_iyi_firma)
            en_ucuz_fiyatlar.append(en_iyi_fiyat_metni)
            
        sonuc_df['En Uygun Tedarikçi'] = en_ucuz_tedarikciler
        sonuc_df['En Uygun Fiyat'] = en_ucuz_fiyatlar
        
        st.write("### Analiz Önizleme")
        st.dataframe(sonuc_df)

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
                if col_num >= last_col - 1:
                    worksheet.set_column(col_num, col_num, 20)
                else:
                    worksheet.set_column(col_num, col_num, max(len(str(value)), 15))

            worksheet.conditional_format(1, 0, row_count, last_col - 2, {'type': 'no_errors', 'format': border_fmt})
            worksheet.conditional_format(1, last_col - 1, row_count, last_col, {'type': 'no_errors', 'format': result_fmt})
            worksheet.freeze_panes(1, 0)

        st.download_button(
            label="📥 Raporu Excel Olarak İndir",
            data=output.getvalue(),
            file_name="Fiyat_Karsilastirma_Raporu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )