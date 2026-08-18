import streamlit as st
import pandas as pd
import re
import io

# 1. JUDUL & DESKRIPSI
st.set_page_config(page_title="Rekonsiliasi Bank Fastpay", layout="wide")
st.title("📊 Rekonsiliasi Bank Fastpay")
st.write("Sistem otomatis mencocokkan data deposit internal dengan mutasi uang masuk di bank.")

st.divider()

# ==========================================
# Inisialisasi Session State (Agar hasil tidak hilang)
# ==========================================
if 'sudah_diproses' not in st.session_state:
    st.session_state.sudah_diproses = False
if 'df_matched' not in st.session_state:
    st.session_state.df_matched = pd.DataFrame()
if 'df_selisih_int' not in st.session_state:
    st.session_state.df_selisih_int = pd.DataFrame()
if 'df_selisih_bnk' not in st.session_state:
    st.session_state.df_selisih_bnk = pd.DataFrame()
if 'pilihan_bank_terakhir' not in st.session_state:
    st.session_state.pilihan_bank_terakhir = ""

# ==========================================
# 2. FITUR PEMILIH BANK & UPLOAD
# ==========================================
st.subheader("1. Pengaturan Data")

opsi_bank = ["", "BRIVA", "BNIVA", "BCAVA", "MANDIRIVA", "BSIVA", "MuamalatVA"]
pilihan_bank = st.selectbox("Pilih Bank Sumber Mutasi:", opsi_bank)

# Reset state jika bank diubah
if pilihan_bank != st.session_state.pilihan_bank_terakhir:
    st.session_state.sudah_diproses = False
    st.session_state.pilihan_bank_terakhir = pilihan_bank

st.subheader("2. Unggah File")
col1, col2 = st.columns(2)

with col1:
    file_int = st.file_uploader("Unggah CSV Dari FMSS", type=['csv', 'xlsx'], key='int')
    
with col2:
    label_bank = f"Unggah CSV Mutasi Bank ({pilihan_bank})" if pilihan_bank != "" else "Unggah CSV Mutasi Bank"
    file_bnk = st.file_uploader(label_bank, type=['csv', 'xlsx'], key='bnk')

# ==========================================
# 3. LOGIKA CROSCEK BERDASARKAN BANK
# ==========================================
if pilihan_bank != "" and file_int and file_bnk:
    st.divider()
    
    # Tombol Proses Utama
    if st.button(f"🚀 Mulai Croscek Data {pilihan_bank}", type="primary"):
        st.session_state.sudah_diproses = True 
        
        # --- BLOK LOGIKA BRIVA ---
        if pilihan_bank == "BRIVA":
            try:
                with st.spinner('Sedang memproses rekonsiliasi BRIVA...'):
                    # Membaca data
                    file_int.seek(0)
                    file_bnk.seek(0)
                    df_int = pd.read_csv(file_int, sep=None, engine='python') if file_int.name.endswith('.csv') else pd.read_excel(file_int)
                    df_bnk = pd.read_csv(file_bnk, sep=None, engine='python') if file_bnk.name.endswith('.csv') else pd.read_excel(file_bnk)

                    # A. Filter Bank (Uang Masuk / MUTASI_KREDIT > 0)
                    df_bnk['MUTASI_KREDIT_NUM'] = pd.to_numeric(df_bnk['MUTASI_KREDIT'], errors='coerce').fillna(0)
                    df_bnk_masuk = df_bnk[df_bnk['MUTASI_KREDIT_NUM'] > 0].copy()

                    # B. Ekstraksi BRIVA (Wajib Awalan 57888)
                    def extract_fastpay(text):
                        if pd.isna(text): return None
                        match = re.search(r'(57888\d{5,15})', str(text))
                        return match.group(1) if match else None

                    df_int['BRIVA_CLEAN'] = df_int['keterangan'].apply(extract_fastpay)
                    df_bnk_masuk['BRIVA_CLEAN'] = df_bnk_masuk['DESK_TRAN'].apply(extract_fastpay)

                    # Buang data yang tidak memiliki kode awalan 57888
                    df_int_final = df_int[df_int['BRIVA_CLEAN'].notna()].copy()
                    df_bnk_final = df_bnk_masuk[df_bnk_masuk['BRIVA_CLEAN'].notna()].copy()

                    # C. Rumus Mutlak Penyesuaian Nominal (+ Rp1.000)
                    df_int_final['NOMINAL_INT_ADJ'] = pd.to_numeric(df_int_final['nominal'], errors='coerce').fillna(0) + 1000
                    
                    # D. Metode Coret 1-lawan-1 (1-to-1 Tallying)
                    list_internal = df_int_final.to_dict('records')
                    list_bank = df_bnk_final.to_dict('records')

                    matched = []
                    unmatched_internal = []
                    unmatched_bank = list_bank.copy()

                    for int_row in list_internal:
                        is_matched = False
                        for i, bank_row in enumerate(unmatched_bank):
                            if (int_row['BRIVA_CLEAN'] == bank_row['BRIVA_CLEAN'] and 
                                int_row['NOMINAL_INT_ADJ'] == bank_row['MUTASI_KREDIT_NUM']):
                                
                                match_record = int_row.copy()
                                match_record['MATCH_MUTASI_KREDIT'] = bank_row['MUTASI_KREDIT_NUM']
                                match_record['MATCH_DESK_TRAN'] = bank_row.get('DESK_TRAN', '')
                                
                                matched.append(match_record)
                                unmatched_bank.pop(i) # Data dicoret
                                is_matched = True
                                break
                                
                        if not is_matched:
                            unmatched_internal.append(int_row)

                    # Simpan hasil ke session state
                    st.session_state.df_matched = pd.DataFrame(matched)
                    st.session_state.df_selisih_int = pd.DataFrame(unmatched_internal)
                    st.session_state.df_selisih_bnk = pd.DataFrame(unmatched_bank)
                    
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis: {e}")
                st.session_state.sudah_diproses = False
                
        # --- BLOK LOGIKA BANK LAIN ---
        else:
            st.warning(f"🚧 Modul rekonsiliasi untuk bank **{pilihan_bank}** sedang dalam tahap persiapan.")
            st.session_state.sudah_diproses = False


    # ==========================================
    # 4. TAMPILAN HASIL DI WEB
    # ==========================================
    if st.session_state.sudah_diproses:
        df_matched = st.session_state.df_matched
        df_selisih_int = st.session_state.df_selisih_int
        df_selisih_bnk = st.session_state.df_selisih_bnk
        
        st.subheader(f"🎯 Ringkasan Rekonsiliasi {pilihan_bank}")
        m1, m2, m3 = st.columns(3)
        m1.metric("✅ Matched Sempurna", f"{len(df_matched)} Trx")
        m2.metric("⚠️ Issue FMSS", f"{len(df_selisih_int)} Trx")
        m3.metric("⚠️ Issue Bank", f"{len(df_selisih_bnk)} Trx")
        
        st.divider()
        
        # TAMPILAN TABEL RINCIAN SELISIH LANGSUNG DI LAYAR
        col_tabel1, col_tabel2 = st.columns(2)
        
        with col_tabel1:
            st.markdown("#### 🚨 Rincian Issue FMSS")
            if not df_selisih_int.empty:
                df_tampil_int = df_selisih_int[['BRIVA_CLEAN', 'nominal']].rename(columns={'BRIVA_CLEAN': 'KODE VA', 'nominal': 'NOMINAL'})
                st.dataframe(df_tampil_int, use_container_width=True, hide_index=True)
            else:
                st.success("Tidak ada issue di FMSS. Semua data cocok!")

        with col_tabel2:
            st.markdown("#### 🚨 Rincian Issue Bank")
            if not df_selisih_bnk.empty:
                df_tampil_bnk = df_selisih_bnk[['BRIVA_CLEAN', 'MUTASI_KREDIT_NUM']].rename(columns={'BRIVA_CLEAN': 'KODE VA', 'MUTASI_KREDIT_NUM': 'NOMINAL'})
                st.dataframe(df_tampil_bnk, use_container_width=True, hide_index=True)
            else:
                st.success("Tidak ada issue di Bank. Semua dana memiliki pasangan!")
                
        st.divider()
        
        # Tombol Download Excel Cadangan
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if not df_selisih_int.empty:
                df_selisih_int.drop(columns=['NOMINAL_INT_ADJ'], errors='ignore').to_excel(writer, sheet_name='ISSUE_FMSS', index=False)
            else:
                pd.DataFrame({'Info': ['Bersih! Tidak ada selisih di FMSS']}).to_excel(writer, sheet_name='ISSUE_FMSS', index=False)
            
            if not df_selisih_bnk.empty:
                df_selisih_bnk.drop(columns=['MUTASI_KREDIT_NUM'], errors='ignore').to_excel(writer, sheet_name='ISSUE_BANK', index=False)
            else:
                pd.DataFrame({'Info': ['Bersih! Tidak ada selisih di Bank']}).to_excel(writer, sheet_name='ISSUE_BANK', index=False)
                
            if not df_matched.empty:
                df_matched.drop(columns=['NOMINAL_INT_ADJ'], errors='ignore').to_excel(writer, sheet_name='MATCHED_OK', index=False)
            else:
                pd.DataFrame({'Info': ['Tidak ada data matched']}).to_excel(writer, sheet_name='MATCHED_OK', index=False)
        
        st.download_button(
            label="📥 Download Laporan Lengkap (.xlsx)",
            data=output.getvalue(),
            file_name=f"Laporan_Croscek_{pilihan_bank}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

elif pilihan_bank == "" and (file_int or file_bnk):
    st.info("💡 Silakan pilih **Bank Sumber Mutasi** terlebih dahulu pada dropdown di atas.")
