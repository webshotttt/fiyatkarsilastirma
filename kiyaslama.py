import streamlit as st
import pandas as pd
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

def temizle_ve_sayiya_cevir(deger):
    if pd.isna(deger): return None
    deger_str = str(deger).replace(',', '.')
    temiz = re.sub(r'[^\d.]', '', deger_str)
    try: return float(temiz)
    except ValueError: return None

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    tedarikci_sutunlari = [col for col in df.columns if any(x in str(col).upper() for x in ['- USD', '- EUR', '- TL'])]
    
    if not tedarikci_sutunlari:
        st.warning("Hata: Uygun formatta tedarikçi sütunu bulunamadı.")
    else:
        # Hesaplama işlemleri
        sonuc_df = df.copy()
        en_ucuz_tedarikciler = []
        en_ucuz_fiyatlar = []
        
        for index, row in sonuc_df.iterrows():
            min_tl = float('inf')
            firma, fiyat_metni = "-", "-"
            
            for col in tedarikci_sutunlari:
                sayi = temizle_ve_sayiya_cevir(row[col])
                if sayi is not None:
                    kur = usd_kur if 'USD' in col.upper() else (eur_kur if 'EUR' in col.upper() else 1)
                    birim = "USD" if 'USD' in col.upper() else ("EUR" if 'EUR' in col.upper() else "TL")
                    tl_deger = sayi * kur
                    if tl_deger < min_tl:
                        min_tl = tl_deger
                        firma = col.split('-')[0].strip()
                        fiyat_metni = f"{sayi} {birim}"
            
            en_ucuz_tedarikciler.append(firma)
            en_ucuz_fiyatlar.append(fiyat_metni)
            
        sonuc_df['En Uygun Tedarikçi'] = en_ucuz_tedarikciler
        sonuc_df['En Uygun Fiyat'] = en_ucuz_fiyatlar
        
        st.write("### Analiz Önizleme")
        st.dataframe(sonuc_df)

        # --- Gelişmiş Excel Tasarımı Bölümü ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sonuc_df.to_excel(writer, index=False, sheet_name='Analiz Raporu')
            
            workbook  = writer.book
            worksheet = writer.sheets['Analiz Raporu']

            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'fg_color': '#1F4E78',
                'font_color': 'white',
                'border': 1
            })

            result_format = workbook.add_format({
                'bold': True,
                'fg_color': '#FFEB9C',
                'border': 1
            })

            cell_format = workbook.add_format({'border': 1})

            last_col = len(sonuc_df.columns) - 1

            # Başlıkları formatla ve sütun genişliklerini ayarla
            for col_num, value in enumerate(sonuc_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                
                # Son iki sütun için sabit 20, diğerleri için minimum 15 genişlik ayarı
                if col_num >= last_col - 1:
                    worksheet.set_column(col_num, col_num, 20)
                else:
                    column_len = max(len(str(value)), 15)
                    worksheet.set_column(col_num, col_num, column_len)

            # SADECE dolu olan satırlara (len(sonuc_df) kadar) format ve kenarlık uygula
            
            # 1. Normal sütunlar için sadece kenarlık
            worksheet.conditional_format(1, 0, len(sonuc_df), last_col - 2, {
                'type': 'no_errors',
                'format': cell_format
            })

            # 2. Son iki sütun için sarı arka plan ve kenarlık
            worksheet.conditional_format(1, last_col - 1, len(sonuc_df), last_col, {
                'type': 'no_errors',
                'format': result_format
            })

            worksheet.freeze_panes(1, 0)

        st.download_button(
            label="📥 Raporu Excel Olarak İndir",
            data=output.getvalue(),
            file_name="Fiyat_Karsilastirma_Raporu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )