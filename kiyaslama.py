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
st.markdown("Kurlar **TCMB** üzerinden canlı çekilmektedir. Analiz raporuna **Dinamik Kıyaslamalı Analiz** sütunu eklenmiştir.")

# TCMB Kur Çekme Motoru
@st.cache_data(ttl=60)
def canli_kurlari_cek():
    try:
        url = "https://www.tcmb.gov.tr/kurlar/today.xml"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            xml_data = response.read()
        root = ET.fromstring(xml_data)
        usd_kur = float(root.find(".//Currency[@CurrencyCode='USD']/ForexSelling").text)
        eur_kur = float(root.find(".//Currency[@CurrencyCode='EUR']/ForexSelling").text)
        guncelleme_zamani = datetime.now().strftime("%d.%m.%Y - %H:%M")
        return round(usd_kur, 4), round(eur_kur, 4), guncelleme_zamani
    except Exception:
        return None, None, None

canli_usd, canli_eur, son_guncelleme = canli_kurlari_cek()

if canli_usd and canli_eur:
    st.success(f"✅ Kurlar TCMB'den Alındı. ({son_guncelleme})")
    default_usd, default_eur = float(canli_usd), float(canli_eur)
else:
    st.warning("⚠️ Canlı kura ulaşılamadı, manuel giriş yapınız.")
    default_usd, default_eur = 44.00, 52.00

col1, col2 = st.columns(2)
usd_kur = col1.number_input("Güncel USD Kuru (TL)", value=default_usd, step=0.01, format="%.4f")
eur_kur = col2.number_input("Güncel EUR Kuru (TL)", value=default_eur, step=0.01, format="%.4f")

uploaded_file = st.file_uploader("Excel Dosyasını Yükleyiniz", type=["xlsx", "xls"])

if uploaded_file:
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active 
    
    tedarikci_sutunlari = df.columns[6:].tolist()
    
    if not tedarikci_sutunlari:
        st.warning("Hata: Tedarikçi sütunu bulunamadı.")
    else:
        sonuc_df = df.copy()
        en_ucuz_tedarikciler, en_ucuz_fiyatlar, kiyaslama_notlari = [], [], []
        
        for index, row in sonuc_df.iterrows():
            teklifler = []
            
            excel_row = index + 2 
            for col_name in tedarikci_sutunlari:
                excel_col = df.columns.get_loc(col_name) + 1
                cell = ws.cell(row=excel_row, column=excel_col)
                deger = cell.value
                
                if deger is not None and str(deger).strip() != "":
                    if isinstance(deger, (int, float)): sayi = float(deger)
                    else:
                        match = re.search(r"(\d+(?:\.\d+)?)", str(deger).replace(',', '.'))
                        sayi = float(match.group(1)) if match else None
                    
                    if sayi is not None:
                        fmt, txt = str(cell.number_format).upper(), str(deger).upper()
                        birim = "EUR" if any(x in f"{fmt} {txt}" for x in ["€", "EUR"]) else \
                                ("USD" if any(x in f"{fmt} {txt}" for x in ["$", "USD"]) else "TL")
                        
                        kur = usd_kur if birim == "USD" else (eur_kur if birim == "EUR" else 1)
                        tl_deger = sayi * kur
                        
                        teklifler.append({
                            'tl': tl_deger,
                            'firma': col_name,
                            'sayi': sayi,
                            'birim': birim
                        })

            if teklifler:
                # Teklifleri TL bazında küçükten büyüğe sırala
                teklifler.sort(key=lambda x: x['tl'])
                en_ucuz = teklifler[0]
                
                en_ucuz_tedarikciler.append(en_ucuz['firma'])
                en_ucuz_fiyatlar.append(f"{en_ucuz['sayi']} {en_ucuz['birim']}")
                
                # Kıyaslama Mantığı (Aynı Kur / Farklı Kur)
                if len(teklifler) > 1:
                    ikinci = teklifler[1]
                    
                    if en_ucuz['birim'] == ikinci['birim']:
                        # İki firma da aynı para biriminde teklif vermişse
                        fark = ikinci['sayi'] - en_ucuz['sayi']
                        kullanilacak_birim = en_ucuz['birim']
                    else:
                        # Firmalar farklı para birimlerinde teklif vermişse (USD'ye çevir)
                        fark_tl = ikinci['tl'] - en_ucuz['tl']
                        fark = fark_tl / usd_kur
                        kullanilacak_birim = "USD"
                    
                    if fark > 0:
                        kiyaslama_notlari.append(f"{en_ucuz['firma'].upper()}, {ikinci['firma'].upper()}'dan {fark:.2f} {kullanilacak_birim} daha ucuz.")
                    else:
                        kiyaslama_notlari.append(f"{en_ucuz['firma'].upper()} ve {ikinci['firma'].upper()} aynı fiyatı verdi.")
                else:
                    kiyaslama_notlari.append("Alternatif teklif bulunamadı.")
            else:
                en_ucuz_tedarikciler.append("-")
                en_ucuz_fiyatlar.append("-")
                kiyaslama_notlari.append("-")
            
        sonuc_df['En Uygun Tedarikçi'] = en_ucuz_tedarikciler
        sonuc_df['En Uygun Fiyat'] = en_ucuz_fiyatlar
        sonuc_df['Karşılaştırmalı Analiz'] = kiyaslama_notlari
        
        st.write("### Analiz Önizleme")
        st.dataframe(sonuc_df)

        # --- Gelişmiş Excel Tasarımı (Hücre Boyama Dahil) ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sonuc_df.to_excel(writer, index=False, sheet_name='Analiz Raporu')
            workbook = writer.book
            worksheet = writer.sheets['Analiz Raporu']

            # Formatlar
            header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'fg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
            winner_cell_fmt = workbook.add_format({'bg_color': '#A9D08E', 'border': 1}) # Yeşil
            summary_fmt = workbook.add_format({'bold': True, 'fg_color': '#FFEB9C', 'border': 1}) # Sarı
            analysis_fmt = workbook.add_format({'bold': True, 'fg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1}) # Açık Yeşil Metin
            border_fmt = workbook.add_format({'border': 1})

            last_col = len(sonuc_df.columns) - 1
            row_count = len(sonuc_df)

            for col_num, value in enumerate(sonuc_df.columns.values):
                worksheet.write(0, col_num, value, header_fmt)
                if col_num == last_col: 
                    worksheet.set_column(col_num, col_num, 50) # Analiz sütununu daha da genişlettim
                else:
                    worksheet.set_column(col_num, col_num, 18)

            for row_idx in range(row_count):
                kazanan_firma = sonuc_df.iloc[row_idx]['En Uygun Tedarikçi']
                for col_idx, col_name in enumerate(tedarikci_sutunlari):
                    excel_col_idx = 6 + col_idx
                    val = sonuc_df.iloc[row_idx][col_name]
                    if pd.isna(val): val = ""
                    
                    if col_name == kazanan_firma:
                        worksheet.write(row_idx + 1, excel_col_idx, val, winner_cell_fmt)
                    else:
                        worksheet.write(row_idx + 1, excel_col_idx, val, border_fmt)

            # Özet sütunları formatı
            worksheet.conditional_format(1, last_col - 2, row_count, last_col - 1, {'type': 'no_errors', 'format': summary_fmt})
            worksheet.conditional_format(1, last_col, row_count, last_col, {'type': 'no_errors', 'format': analysis_fmt})
            
            worksheet.freeze_panes(1, 0)

        st.download_button(
            label="📥 Raporu Excel Olarak İndir",
            data=output.getvalue(),
            file_name="Gemi_Tedarik_Fiyat_Analizi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )