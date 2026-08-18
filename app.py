import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Rekonsiliasi BRIVA Fastpay", layout="wide")
st.title("📊 Web Tracker Rekonsiliasi BRIVA Fastpay")
st.write("Sistem otomatis mencocokkan data deposit internal dengan mutasi uang masuk BRIVA (Awalan 57888).")

# ==========================================
# 1. TAMPILAN UPLOAD FILE
# ==========================================
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. File Panel Internal")
    file_int = st.file_uploader("Unggah CSV Internal (Status Sukses)", type=['csv', 'xlsx'], key='int')
with col2:
    st.subheader("2. File Mutasi Bank")
    file_bnk = st.file_uploader("Unggah CSV Mutasi BRIVA", type=['csv', 'xlsx'], key='bnk')

# ==========================================
# 2. LOGIKA CROSCEK DATA
# ==========================================
if file_int and file_bnk:
    st.divider()
    
    try:
        # Membaca data
        df_int = pd.read_csv(file_int, sep=None, engine='python') if file_int.name.endswith('.csv') else pd.read_excel(file_int)
        df_bnk = pd.read_csv(file_bnk, sep=None, engine='python') if file_bnk.name.endswith('.csv') else pd.read_excel(file_bnk)
        
        if st.button("🚀 Mulai Croscek Data", type="primary"):
            with st.spinner('Sedang memproses dan mencoret data...'):
                
                # A. Filter Status Internal (Sukses) & Bank (Uang Masuk)
                df_int_sukses = df_int[df_int['status'].astype(str).str.upper() == 'SUKSES'].copy()
                df_bnk['MUTASI_KREDIT_NUM'] = pd.to_numeric(df_bnk['MUTASI_KREDIT'], errors='coerce').fillna(0)
                df_bnk_masuk = df_bnk[df_bnk['MUTASI_KREDIT_NUM'] > 0].copy()

                # B. Ekstraksi BRIVA (Wajib Awalan 57888)
                def extract_fastpay(text):
                    if pd.isna(text): return None
                    match = re.search(r'(57888\d{5,15})', str(text))
                    return match.group(1) if match else None

                df_int_sukses['BRIVA_CLEAN'] = df_int_sukses['keterangan'].apply(extract_fastpay)
                df_bnk_masuk['BRIVA_CLEAN'] = df_bnk_masuk['DESK_TRAN'].apply(extract_fastpay)

                # Buang data yang bukan 57888
                df_int_final = df_int_sukses[df_int_sukses['BRIVA_CLEAN'].notna()].copy()
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

                df_matched = pd.DataFrame(matched)
                df_selisih_int = pd.DataFrame(unmatched_internal)
                df_selisih_bnk = pd.DataFrame(unmatched_bank)

                # ==========================================
                # 3. TAMPILAN HASIL DI WEB
                # ==========================================
                st.subheader("🎯 Ringkasan Rekonsiliasi")
                m1, m2, m3 = st.columns(3)
                m1.metric("✅ Matched Sempurna", f"{len(df_matched)} Trx")
                m2.metric("⚠️ Issue Internal", f"{len(df_selisih_int)} Trx")
                m3.metric("⚠️ Issue Bank", f"{len(df_selisih_bnk)} Trx")
                
                # Membuat File Excel untuk Didownload
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    if not df_selisih_int.empty:
                        df_selisih_int.drop(columns=['NOMINAL_INT_ADJ'], errors='ignore').to_excel(writer, sheet_name='ISSUE_INTERNAL', index=False)
                    else:
                        pd.DataFrame({'Info': ['Bersih! Tidak ada selisih di Internal']}).to_excel(writer, sheet_name='ISSUE_INTERNAL', index=False)
                    
                    if not df_selisih_bnk.empty:
                        df_selisih_bnk.drop(columns=['MUTASI_KREDIT_NUM'], errors='ignore').to_excel(writer, sheet_name='ISSUE_BANK', index=False)
                    else:
                        pd.DataFrame({'Info': ['Bersih! Tidak ada selisih di Bank']}).to_excel(writer, sheet_name='ISSUE_BANK', index=False)
                        
                    if not df_matched.empty:
                        df_matched.drop(columns=['NOMINAL_INT_ADJ'], errors='ignore').to_excel(writer, sheet_name='MATCHED_OK', index=False)
                    else:
                        pd.DataFrame({'Info': ['Tidak ada data matched']}).to_excel(writer, sheet_name='MATCHED_OK', index=False)
                
                st.success("Selesai! Laporan siap diunduh.")
                
                st.download_button(
                    label="📥 Download File Laporan Excel (.xlsx)",
                    data=output.getvalue(),
                    file_name="Laporan_Croscek_BRIVA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
    except Exception as e:
        st.error("Terjadi kesalahan saat membaca fail. Pastikan kolom 'keterangan', 'nominal', 'status' ada di panel internal, dan 'DESK_TRAN', 'MUTASI_KREDIT' ada di mutasi bank.")
