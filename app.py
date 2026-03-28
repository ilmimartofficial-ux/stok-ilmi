import streamlit as st
import pandas as pd
from io import BytesIO

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="ILMI STOK SYNC", page_icon="📦")

# --- TEMA WARNA (KUNING & HITAM) ---
st.markdown("""
    <style>
    .main { background-color: #FFFFFF; }
    .stButton>button {
        background-color: #FBC02D;
        color: black;
        border-radius: 10px;
        border: none;
        height: 3em;
        width: 100%;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #FFD54F;
        color: black;
    }
    h1 { color: #000000; text-align: center; background-color: #FBC02D; padding: 20px; border_radius: 10px; }
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

st.write("<h1>📦 ILMI STOK MONITOR</h1>", unsafe_allow_html=True)
st.write("---")

# --- FUNGSI MEMBERSIHKAN EXCEL ---
def auto_clean_excel(file):
    try:
        raw_df = pd.read_excel(file, header=None)
        for i, row in raw_df.iterrows():
            if row.astype(str).str.contains('Kode', case=False).any():
                df = pd.read_excel(file, skiprows=i+1)
                df.columns = raw_df.iloc[i]
                return df.dropna(how='all', axis=1)
    except:
        return None
    return None

def find_col(df, keyword):
    for col in df.columns:
        if keyword.lower() in str(col).lower():
            return col
    return None

# --- UI INPUT ---
st.subheader("1. Upload Data iPos")
file4 = st.file_uploader("Pilih Data iPos 4 (MART)", type=['xlsx'])
file5 = st.file_uploader("Pilih Data iPos 5 (GROSIR)", type=['xlsx'])

st.subheader("2. Pengaturan")
min_stok = st.number_input("Tampilkan Stok IG di bawah angka:", value=10)

# --- PROSES DATA ---
if st.button("🚀 GENERATE DAFTAR ORDER"):
    if file4 and file5:
        with st.spinner('Menghitung stok...'):
            df4 = auto_clean_excel(file4)
            df5 = auto_clean_excel(file5)

            if df4 is not None and df5 is not None:
                c_kode_ig = find_col(df5, 'Kode')
                c_nama_ig = find_col(df5, 'Nama')
                c_stok_ig = find_col(df5, 'Stok')
                c_modal_ig = find_col(df5, 'Pokok')
                
                c_kode_im = find_col(df4, 'Kode')
                c_stok_im = find_col(df4, 'Stok')

                # Filter & Sinkron
                ig = df5[[c_kode_ig, c_nama_ig, c_stok_ig, c_modal_ig]].copy()
                ig.columns = ['KODE', 'NAMA_ITEM', 'IG', 'MODAL']
                
                im = df4[[c_kode_im, c_stok_im]].copy()
                im.columns = ['KODE', 'IM']

                final = pd.merge(ig, im, on='KODE', how='left').fillna(0)
                final['IG'] = pd.to_numeric(final['IG'], errors='coerce').fillna(0)
                final = final[final['IG'] < min_stok].copy()

                # Kolom Kosong
                final['ORDER'] = "" 
                final['SUPLIER'] = "" 
                
                order = ['KODE', 'NAMA_ITEM', 'IG', 'IM', 'MODAL', 'ORDER', 'SUPLIER']
                final = final[order]

                # Tampilkan Preview di Web
                st.success(f"Berhasil! Ditemukan {len(final)} barang kritis.")
                st.dataframe(final)

                # Tombol Download Hasil
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 DOWNLOAD EXCEL DAFTAR ORDER",
                    data=output.getvalue(),
                    file_name="Daftar_Order_Ilmi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Kolom 'Kode' tidak ditemukan. Cek format Excel Anda.")
    else:
        st.warning("Silakan upload kedua file terlebih dahulu.")

st.caption("v8.0 - Khusus Web & Mobile Android")