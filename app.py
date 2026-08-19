import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Rekonsiliasi Bank Fastpay",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Rekonsiliasi Bank Fastpay")
st.write(
    "Dashboard rekonsiliasi otomatis antara data deposit FMSS "
    "dengan mutasi bank."
)

st.divider()


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_STATE = {
    "sudah_diproses": False,
    "df_matched": pd.DataFrame(),
    "df_selisih_int": pd.DataFrame(),
    "df_selisih_bnk": pd.DataFrame(),
    "df_invalid_int": pd.DataFrame(),
    "df_invalid_bnk": pd.DataFrame(),
    "recon_dates": [],
    "summary": {},
    "pilihan_bank_terakhir": ""
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def find_column(df, candidates, required=True):
    """
    Mencari nama kolom berdasarkan exact match / case-insensitive.
    """

    if df is None or df.empty:
        if required:
            raise ValueError(
                f"Data kosong. Tidak dapat mencari kolom: {candidates}"
            )
        return None

    # Exact
    for candidate in candidates:
        if candidate in df.columns:
            return candidate

    # Case insensitive
    mapping = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    for candidate in candidates:
        key = str(candidate).strip().lower()
        if key in mapping:
            return mapping[key]

    if required:
        raise ValueError(
            f"Kolom tidak ditemukan. Dicari salah satu dari: {candidates}. "
            f"Kolom tersedia: {list(df.columns)}"
        )

    return None


def read_uploaded_file(uploaded_file):
    """
    Membaca CSV / XLSX secara aman.
    """

    uploaded_file.seek(0)

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        return pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python"
        )

    elif filename.endswith(".xlsx"):
        return pd.read_excel(uploaded_file)

    else:
        raise ValueError(
            f"Format file tidak didukung: {uploaded_file.name}"
        )


def clean_numeric(series):
    """
    Normalisasi nominal menjadi numeric.
    Menangani format angka umum.
    """

    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    cleaned = (
        series.astype(str)
        .str.replace("Rp", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", "", regex=False)
    )

    return pd.to_numeric(
        cleaned,
        errors="coerce"
    ).fillna(0)


def parse_datetime(series):
    """
    Parsing tanggal/waktu secara aman.
    """

    return pd.to_datetime(
        series,
        errors="coerce"
    )


def extract_va(text):
    """
    Mengambil VA Fastpay / Rajabiller.

    57888 = BRIVA Fastpay
    57708 = BRIVA Rajabiller
    """

    if pd.isna(text):
        return None

    text = str(text)

    # Cari 57888 atau 57708 + 5 sampai 15 digit berikutnya
    match = re.search(
        r"(57(?:888|708)\d{5,15})",
        text
    )

    if match:
        return match.group(1)

    return None


def classify_va(va):
    if pd.isna(va) or va is None or str(va).strip() == "":
        return "INVALID VA"

    va = str(va)

    if va.startswith("57888"):
        return "BRIVA FASTPAY"

    if va.startswith("57708"):
        return "BRIVA RAJABILLER"

    return "UNKNOWN"


def classify_bank_transaction(description):
    """
    Klasifikasi sederhana transaksi bank.
    Tidak digunakan sebagai syarat matching.
    """

    text = str(description).upper()

    if "ATM" in text:
        return "ATM / MANUAL"

    if "TRF BERSAMA" in text:
        return "TRANSFER / MANUAL"

    if "BRIVA" in text:
        return "BRIVA"

    if "BFVA" in text:
        return "BFVA"

    if "VA" in text:
        return "VA"

    return "OTHER"


def safe_date_string(dates):
    if not dates:
        return "-"

    sorted_dates = sorted(dates)

    if len(sorted_dates) == 1:
        return sorted_dates[0].strftime("%d %B %Y")

    return (
        f"{sorted_dates[0].strftime('%d %B %Y')} "
        f"s/d {sorted_dates[-1].strftime('%d %B %Y')}"
    )


def format_rupiah(value):
    try:
        value = float(value)
    except:
        value = 0

    return "Rp {:,.0f}".format(value).replace(",", ".")


def classify_issue_bank(description):
    """
    Klasifikasi issue Bank Only agar lebih mudah diinvestigasi.
    """

    category = classify_bank_transaction(description)

    if category in ["ATM / MANUAL", "TRANSFER / MANUAL"]:
        return "BANK_ONLY - MANUAL/ATM"

    if category == "BRIVA":
        return "BANK_ONLY - BRIVA"

    if category == "BFVA":
        return "BANK_ONLY - BFVA"

    return "BANK_ONLY - OTHER"


# ============================================================
# UI - PILIH BANK
# ============================================================

st.subheader("1. Pengaturan Data")

opsi_bank = [
    "",
    "BRIVA",
    "BNIVA",
    "BCAVA",
    "MANDIRIVA",
    "BSIVA",
    "MuamalatVA"
]

pilihan_bank = st.selectbox(
    "Pilih Bank Sumber Mutasi:",
    opsi_bank
)


# Reset hasil jika bank berubah
if pilihan_bank != st.session_state.pilihan_bank_terakhir:

    st.session_state.sudah_diproses = False

    st.session_state.df_matched = pd.DataFrame()
    st.session_state.df_selisih_int = pd.DataFrame()
    st.session_state.df_selisih_bnk = pd.DataFrame()
    st.session_state.df_invalid_int = pd.DataFrame()
    st.session_state.df_invalid_bnk = pd.DataFrame()
    st.session_state.recon_dates = []
    st.session_state.summary = {}

    st.session_state.pilihan_bank_terakhir = pilihan_bank


# ============================================================
# UI - UPLOAD
# ============================================================

st.subheader("2. Unggah File")

if pilihan_bank == "BRIVA":

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("### 📄 FMSS")
        file_int = st.file_uploader(
            "Upload data FMSS",
            type=["csv", "xlsx"],
            key="fmss_briva"
        )

    with col2:
        st.markdown("### 🏦 BRIVA Fastpay — 57888")
        file_bnk_57888 = st.file_uploader(
            "Upload mutasi BRIVA 57888",
            type=["csv", "xlsx"],
            key="briva_57888"
        )

    with col3:
        st.markdown("### 🏦 BRIVA Rajabiller — 57708")
        file_bnk_57708 = st.file_uploader(
            "Upload mutasi BRIVA 57708",
            type=["csv", "xlsx"],
            key="briva_57708"
        )

else:

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 📄 FMSS")
        file_int = st.file_uploader(
            "Upload data FMSS",
            type=["csv", "xlsx"],
            key="fmss_general"
        )

    with col2:
        st.markdown(f"### 🏦 Mutasi {pilihan_bank}")
        file_bnk_general = st.file_uploader(
            f"Upload mutasi {pilihan_bank}",
            type=["csv", "xlsx"],
            key="bank_general"
        )

    file_bnk_57888 = None
    file_bnk_57708 = None


# ============================================================
# KONFIGURASI FEE BRIVA
# ============================================================

if pilihan_bank == "BRIVA":

    st.subheader("3. Konfigurasi Fee")

    col_fee1, col_fee2 = st.columns(2)

    with col_fee1:
        fee_57888 = st.number_input(
            "Fee Fastpay (57888)",
            min_value=0,
            value=1000,
            step=100,
            format="%d"
        )

    with col_fee2:
        fee_57708 = st.number_input(
            "Fee Rajabiller (57708)",
            min_value=0,
            value=1000,
            step=100,
            format="%d"
        )

    st.caption(
        "Rumus pencocokan: Nominal FMSS + Fee = Nominal mutasi bank."
    )

else:

    fee_57888 = 1000
    fee_57708 = 1000


# ============================================================
# BUTTON PROSES
# ============================================================

can_process = False

if pilihan_bank == "BRIVA":

    if file_int and file_bnk_57888 and file_bnk_57708:
        can_process = True

else:

    if file_int and file_bnk_general:
        can_process = True


if can_process:

    st.divider()

    if st.button(
        f"🚀 Mulai Croscek Data {pilihan_bank}",
        type="primary",
        use_container_width=True
    ):

        st.session_state.sudah_diproses = False

        try:

            with st.spinner(
                "Sedang memproses rekonsiliasi..."
            ):

                # ====================================================
                # LOAD FMSS
                # ====================================================

                df_int = read_uploaded_file(file_int)

                # Cari kolom FMSS
                col_status = find_column(
                    df_int,
                    ["status", "STATUS"]
                )

                col_keterangan_int = find_column(
                    df_int,
                    [
                        "keterangan",
                        "KETERANGAN",
                        "description",
                        "DESKRIPSI"
                    ]
                )

                col_nominal_int = find_column(
                    df_int,
                    [
                        "nominal",
                        "NOMINAL",
                        "amount",
                        "AMOUNT"
                    ]
                )

                col_tanggal_int = find_column(
                    df_int,
                    [
                        "tanggal_transfer",
                        "TANGGAL_TRANSFER",
                        "tanggal",
                        "TANGGAL",
                        "tgl_transfer",
                        "TGL_TRANSFER"
                    ]
                )

                # ====================================================
                # FILTER FMSS SUKSES
                # ====================================================

                df_int = df_int.copy()

                df_int["_STATUS_CLEAN"] = (
                    df_int[col_status]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df_int_sukses = df_int[
                    df_int["_STATUS_CLEAN"] == "SUKSES"
                ].copy()

                # ====================================================
                # TANGGAL REKONSILIASI
                # ====================================================

                df_int_sukses["_TANGGAL_DT"] = parse_datetime(
                    df_int_sukses[col_tanggal_int]
                )

                df_int_sukses = df_int_sukses[
                    df_int_sukses["_TANGGAL_DT"].notna()
                ].copy()

                if df_int_sukses.empty:
                    raise ValueError(
                        "Tidak ada transaksi FMSS SUKSES "
                        "dengan tanggal yang valid."
                    )

                recon_dates = sorted(
                    df_int_sukses["_TANGGAL_DT"]
                    .dt.date
                    .dropna()
                    .unique()
                )

                st.session_state.recon_dates = recon_dates

                # ====================================================
                # EXTRACT VA FMSS
                # ====================================================

                df_int_sukses["KODE_VA"] = (
                    df_int_sukses[col_keterangan_int]
                    .apply(extract_va)
                )

                df_int_sukses["JENIS_VA"] = (
                    df_int_sukses["KODE_VA"]
                    .apply(classify_va)
                )

                # ====================================================
                # INVALID VA FMSS
                # ====================================================

                df_invalid_int = df_int_sukses[
                    df_int_sukses["KODE_VA"].isna()
                ].copy()

                # Hanya transaksi yang punya VA valid yang masuk
                # engine matching
                df_int_valid = df_int_sukses[
                    df_int_sukses["KODE_VA"].notna()
                ].copy()

                # ====================================================
                # NOMINAL FMSS
                # ====================================================

                df_int_valid["NOMINAL_ASLI"] = clean_numeric(
                    df_int_valid[col_nominal_int]
                )

                # ====================================================
                # BANK PROCESSING
                # ====================================================

                bank_sources = []

                if pilihan_bank == "BRIVA":

                    # -----------------------------------------------
                    # BRIVA 57888
                    # -----------------------------------------------

                    df_57888 = read_uploaded_file(
                        file_bnk_57888
                    )

                    col_credit_57888 = find_column(
                        df_57888,
                        [
                            "MUTASI_KREDIT",
                            "mutasi_kredit",
                            "KREDIT",
                            "kredit"
                        ]
                    )

                    col_desc_57888 = find_column(
                        df_57888,
                        [
                            "DESK_TRAN",
                            "desk_tran",
                            "KETERANGAN",
                            "keterangan",
                            "DESCRIPTION",
                            "description"
                        ]
                    )

                    col_date_57888 = find_column(
                        df_57888,
                        [
                            "TGL_TRAN",
                            "tgl_tran",
                            "TANGGAL_TRAN",
                            "tanggal_tran",
                            "TANGGAL",
                            "tanggal"
                        ]
                    )

                    df_57888 = df_57888.copy()

                    df_57888["_BANK_TYPE"] = (
                        df_57888[col_desc_57888]
                        .apply(classify_bank_transaction)
                    )

                    df_57888["_TANGGAL_DT"] = parse_datetime(
                        df_57888[col_date_57888]
                    )

                    df_57888["_CREDIT_NUM"] = clean_numeric(
                        df_57888[col_credit_57888]
                    )

                    # HANYA TRANSAKSI PADA TANGGAL REKONSILIASI
                    df_57888 = df_57888[
                        df_57888["_TANGGAL_DT"]
                        .dt.date
                        .isin(recon_dates)
                    ].copy()

                    # HANYA UANG MASUK
                    df_57888 = df_57888[
                        df_57888["_CREDIT_NUM"] > 0
                    ].copy()

                    df_57888["KODE_VA"] = (
                        df_57888[col_desc_57888]
                        .apply(extract_va)
                    )

                    df_57888["JENIS_VA"] = (
                        df_57888["KODE_VA"]
                        .apply(classify_va)
                    )

                    df_57888["SOURCE_BANK"] = (
                        "BRIVA FASTPAY 57888"
                    )

                    bank_sources.append(df_57888)

                    # -----------------------------------------------
                    # BRIVA 57708
                    # -----------------------------------------------

                    df_57708 = read_uploaded_file(
                        file_bnk_57708
                    )

                    col_credit_57708 = find_column(
                        df_57708,
                        [
                            "MUTASI_KREDIT",
                            "mutasi_kredit",
                            "KREDIT",
                            "kredit"
                        ]
                    )

                    col_desc_57708 = find_column(
                        df_57708,
                        [
                            "DESK_TRAN",
                            "desk_tran",
                            "KETERANGAN",
                            "keterangan",
                            "DESCRIPTION",
                            "description"
                        ]
                    )

                    col_date_57708 = find_column(
                        df_57708,
                        [
                            "TGL_TRAN",
                            "tgl_tran",
                            "TANGGAL_TRAN",
                            "tanggal_tran",
                            "TANGGAL",
                            "tanggal"
                        ]
                    )

                    df_57708 = df_57708.copy()

                    df_57708["_BANK_TYPE"] = (
                        df_57708[col_desc_57708]
                        .apply(classify_bank_transaction)
                    )

                    df_57708["_TANGGAL_DT"] = parse_datetime(
                        df_57708[col_date_57708]
                    )

                    df_57708["_CREDIT_NUM"] = clean_numeric(
                        df_57708[col_credit_57708]
                    )

                    # HANYA TANGGAL REKONSILIASI
                    df_57708 = df_57708[
                        df_57708["_TANGGAL_DT"]
                        .dt.date
                        .isin(recon_dates)
                    ].copy()

                    # HANYA UANG MASUK
                    df_57708 = df_57708[
                        df_57708["_CREDIT_NUM"] > 0
                    ].copy()

                    df_57708["KODE_VA"] = (
                        df_57708[col_desc_57708]
                        .apply(extract_va)
                    )

                    df_57708["JENIS_VA"] = (
                        df_57708["KODE_VA"]
                        .apply(classify_va)
                    )

                    df_57708["SOURCE_BANK"] = (
                        "BRIVA RAJABILLER 57708"
                    )

                    bank_sources.append(df_57708)

                else:

                    # =================================================
                    # BANK LAIN
                    # =================================================

                    df_general = read_uploaded_file(
                        file_bnk_general
                    )

                    col_credit = find_column(
                        df_general,
                        [
                            "MUTASI_KREDIT",
                            "mutasi_kredit",
                            "KREDIT",
                            "kredit"
                        ]
                    )

                    col_desc = find_column(
                        df_general,
                        [
                            "DESK_TRAN",
                            "desk_tran",
                            "KETERANGAN",
                            "keterangan",
                            "DESCRIPTION",
                            "description"
                        ]
                    )

                    col_date = find_column(
                        df_general,
                        [
                            "TGL_TRAN",
                            "tgl_tran",
                            "TANGGAL_TRAN",
                            "tanggal_tran",
                            "TANGGAL",
                            "tanggal"
                        ]
                    )

                    df_general["_BANK_TYPE"] = (
                        df_general[col_desc]
                        .apply(classify_bank_transaction)
                    )

                    df_general["_TANGGAL_DT"] = parse_datetime(
                        df_general[col_date]
                    )

                    df_general["_CREDIT_NUM"] = clean_numeric(
                        df_general[col_credit]
                    )

                    df_general = df_general[
                        df_general["_TANGGAL_DT"]
                        .dt.date
                        .isin(recon_dates)
                    ].copy()

                    df_general = df_general[
                        df_general["_CREDIT_NUM"] > 0
                    ].copy()

                    df_general["KODE_VA"] = (
                        df_general[col_desc]
                        .apply(extract_va)
                    )

                    df_general["JENIS_VA"] = (
                        df_general["KODE_VA"]
                        .apply(classify_va)
                    )

                    df_general["SOURCE_BANK"] = pilihan_bank

                    bank_sources.append(df_general)

                # ====================================================
                # COMBINE BANK DATA
                # ====================================================

                if bank_sources:

                    df_bank = pd.concat(
                        bank_sources,
                        ignore_index=True
                    )

                else:

                    df_bank = pd.DataFrame()

                # ====================================================
                # BANK INVALID VA
                # ====================================================

                df_invalid_bnk = df_bank[
                    df_bank["KODE_VA"].isna()
                ].copy()

                # HANYA VA YANG VALID UNTUK MATCHING
                df_bank_valid = df_bank[
                    df_bank["KODE_VA"].notna()
                ].copy()

                # ====================================================
                # MATCHING ENGINE
                # ====================================================

                # Pisahkan FMSS berdasarkan prefix
                # supaya 57888 tidak pernah dibandingkan
                # dengan 57708.

                df_int_valid["EXPECTED_BANK"] = (
                    df_int_valid["NOMINAL_ASLI"]
                )

                df_int_valid.loc[
                    df_int_valid["JENIS_VA"] == "BRIVA FASTPAY",
                    "EXPECTED_BANK"
                ] = (
                    df_int_valid.loc[
                        df_int_valid["JENIS_VA"] == "BRIVA FASTPAY",
                        "NOMINAL_ASLI"
                    ] + fee_57888
                )

                df_int_valid.loc[
                    df_int_valid["JENIS_VA"] == "BRIVA RAJABILLER",
                    "EXPECTED_BANK"
                ] = (
                    df_int_valid.loc[
                        df_int_valid["JENIS_VA"] == "BRIVA RAJABILLER",
                        "NOMINAL_ASLI"
                    ] + fee_57708
                )

                # ----------------------------------------------------
                # Buat list bank yang bisa dicoret
                # ----------------------------------------------------

                bank_records = df_bank_valid.to_dict(
                    "records"
                )

                matched = []
                unmatched_internal = []

                # ----------------------------------------------------
                # 1-to-1 matching
                # ----------------------------------------------------

                for _, int_row in df_int_valid.iterrows():

                    matched_index = None

                    for i, bank_row in enumerate(bank_records):

                        same_va = (
                            str(int_row["KODE_VA"])
                            == str(bank_row["KODE_VA"])
                        )

                        same_nominal = (
                            float(int_row["EXPECTED_BANK"])
                            == float(bank_row["_CREDIT_NUM"])
                        )

                        if same_va and same_nominal:

                            matched_index = i
                            break

                    if matched_index is not None:

                        bank_row = bank_records.pop(
                            matched_index
                        )

                        record = int_row.to_dict()

                        record["MATCH_MUTASI_KREDIT"] = (
                            bank_row["_CREDIT_NUM"]
                        )

                        record["MATCH_DESK_TRAN"] = (
                            bank_row.get(
                                col_desc_57888
                                if (
                                    pilihan_bank == "BRIVA"
                                    and bank_row.get(
                                        "SOURCE_BANK"
                                    )
                                    == "BRIVA FASTPAY 57888"
                                )
                                else col_desc_57708
                                if (
                                    pilihan_bank == "BRIVA"
                                    and bank_row.get(
                                        "SOURCE_BANK"
                                    )
                                    == "BRIVA RAJABILLER 57708"
                                )
                                else "DESK_TRAN",
                                ""
                            )
                        )

                        record["SOURCE_BANK"] = (
                            bank_row.get(
                                "SOURCE_BANK",
                                pilihan_bank
                            )
                        )

                        record["BANK_TYPE"] = (
                            bank_row.get(
                                "_BANK_TYPE",
                                ""
                            )
                        )

                        record["STATUS_MATCH"] = "MATCHED"

                        matched.append(record)

                    else:

                        record = int_row.to_dict()

                        record["STATUS_MATCH"] = (
                            "FMSS_ONLY"
                        )

                        unmatched_internal.append(
                            record
                        )

                # ====================================================
                # BANK YANG TERSISA = BANK ONLY
                # ====================================================

                unmatched_bank = []

                for bank_row in bank_records:

                    record = bank_row.copy()

                    record["STATUS_MATCH"] = (
                        classify_issue_bank(
                            bank_row.get(
                                "_BANK_TYPE",
                                ""
                            )
                        )
                    )

                    unmatched_bank.append(
                        record
                    )

                # ====================================================
                # DATAFRAME HASIL
                # ====================================================

                df_matched = pd.DataFrame(
                    matched
                )

                df_selisih_int = pd.DataFrame(
                    unmatched_internal
                )

                df_selisih_bnk = pd.DataFrame(
                    unmatched_bank
                )

                # ====================================================
                # SUMMARY
                # ====================================================

                matched_count = len(
                    df_matched
                )

                fmss_only_count = len(
                    df_selisih_int
                )

                bank_only_count = len(
                    df_selisih_bnk
                )

                invalid_int_count = len(
                    df_invalid_int
                )

                invalid_bnk_count = len(
                    df_invalid_bnk
                )

                matched_nominal = (
                    df_matched["NOMINAL_ASLI"].sum()
                    if (
                        not df_matched.empty
                        and "NOMINAL_ASLI"
                        in df_matched.columns
                    )
                    else 0
                )

                fmss_only_nominal = (
                    df_selisih_int["NOMINAL_ASLI"].sum()
                    if (
                        not df_selisih_int.empty
                        and "NOMINAL_ASLI"
                        in df_selisih_int.columns
                    )
                    else 0
                )

                bank_only_nominal = (
                    df_selisih_bnk["_CREDIT_NUM"].sum()
                    if (
                        not df_selisih_bnk.empty
                        and "_CREDIT_NUM"
                        in df_selisih_bnk.columns
                    )
                    else 0
                )

                summary = {
                    "matched_count": matched_count,
                    "fmss_only_count": fmss_only_count,
                    "bank_only_count": bank_only_count,
                    "invalid_int_count": invalid_int_count,
                    "invalid_bnk_count": invalid_bnk_count,
                    "matched_nominal": matched_nominal,
                    "fmss_only_nominal": fmss_only_nominal,
                    "bank_only_nominal": bank_only_nominal
                }

                # ====================================================
                # SAVE SESSION STATE
                # ====================================================

                st.session_state.df_matched = df_matched
                st.session_state.df_selisih_int = (
                    df_selisih_int
                )
                st.session_state.df_selisih_bnk = (
                    df_selisih_bnk
                )
                st.session_state.df_invalid_int = (
                    df_invalid_int
                )
                st.session_state.df_invalid_bnk = (
                    df_invalid_bnk
                )
                st.session_state.summary = summary

                st.session_state.sudah_diproses = True

        except Exception as e:

            st.session_state.sudah_diproses = False

            st.error(
                "❌ Terjadi kesalahan saat memproses data."
            )

            st.exception(e)


# ============================================================
# HASIL REKONSILIASI
# ============================================================

if st.session_state.sudah_diproses:

    df_matched = st.session_state.df_matched
    df_selisih_int = st.session_state.df_selisih_int
    df_selisih_bnk = st.session_state.df_selisih_bnk
    df_invalid_int = st.session_state.df_invalid_int
    df_invalid_bnk = st.session_state.df_invalid_bnk

    summary = st.session_state.summary

    st.divider()

    # ========================================================
    # HEADER
    # ========================================================

    st.subheader(
        f"🎯 Ringkasan Rekonsiliasi {pilihan_bank}"
    )

    st.caption(
        f"Periode rekonsiliasi: "
        f"**{safe_date_string(st.session_state.recon_dates)}**"
    )

    # ========================================================
    # METRIC
    # ========================================================

    m1, m2, m3, m4 = st.columns(4)

    m1.metric(
        "✅ Matched Sempurna",
        f"{summary['matched_count']:,} Trx"
    )

    m2.metric(
        "⚠️ Issue FMSS",
        f"{summary['fmss_only_count']:,} Trx"
    )

    m3.metric(
        "⚠️ Issue Bank",
        f"{summary['bank_only_count']:,} Trx"
    )

    m4.metric(
        "🚨 Invalid VA",
        f"{summary['invalid_int_count'] + summary['invalid_bnk_count']:,} Trx"
    )

    # ========================================================
    # MATCH RATE
    # ========================================================

    total_fmss_valid = (
        len(df_matched)
        + len(df_selisih_int)
    )

    total_bank_valid = (
        len(df_matched)
        + len(df_selisih_bnk)
    )

    fmss_match_rate = (
        len(df_matched)
        / total_fmss_valid
        * 100
        if total_fmss_valid > 0
        else 0
    )

    bank_match_rate = (
        len(df_matched)
        / total_bank_valid
        * 100
        if total_bank_valid > 0
        else 0
    )

    st.divider()

    r1, r2 = st.columns(2)

    r1.metric(
        "📈 Match Rate FMSS",
        f"{fmss_match_rate:.4f}%"
    )

    r2.metric(
        "📈 Match Rate Bank",
        f"{bank_match_rate:.4f}%"
    )

    # ========================================================
    # NOMINAL SUMMARY
    # ========================================================

    st.subheader("💰 Ringkasan Nominal")

    n1, n2, n3 = st.columns(3)

    n1.metric(
        "Matched",
        format_rupiah(
            summary["matched_nominal"]
        )
    )

    n2.metric(
        "Issue FMSS",
        format_rupiah(
            summary["fmss_only_nominal"]
        )
    )

    n3.metric(
        "Issue Bank",
        format_rupiah(
            summary["bank_only_nominal"]
        )
    )

    # ========================================================
    # ISSUE FMSS / BANK
    # ========================================================

    st.divider()

    col_issue1, col_issue2 = st.columns(2)

    # ========================================================
    # ISSUE FMSS
    # ========================================================

    with col_issue1:

        st.subheader("🚨 Issue FMSS")

        if not df_selisih_int.empty:

            display_int = pd.DataFrame()

            display_int["KODE VA"] = (
                df_selisih_int["KODE_VA"]
            )

            display_int["JENIS VA"] = (
                df_selisih_int["JENIS_VA"]
            )

            display_int["NOMINAL"] = (
                df_selisih_int["NOMINAL_ASLI"]
            )

            display_int["EXPECTED BANK"] = (
                df_selisih_int["EXPECTED_BANK"]
            )

            display_int["ISSUE"] = "FMSS_ONLY"

            st.dataframe(
                display_int,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue FMSS. "
                "Semua transaksi FMSS memiliki pasangan bank."
            )

    # ========================================================
    # ISSUE BANK
    # ========================================================

    with col_issue2:

        st.subheader("🚨 Issue Bank")

        if not df_selisih_bnk.empty:

            display_bnk = pd.DataFrame()

            display_bnk["KODE VA"] = (
                df_selisih_bnk["KODE_VA"]
            )

            display_bnk["JENIS VA"] = (
                df_selisih_bnk["JENIS_VA"]
            )

            display_bnk["NOMINAL"] = (
                df_selisih_bnk["_CREDIT_NUM"]
            )

            display_bnk["SOURCE BANK"] = (
                df_selisih_bnk["SOURCE_BANK"]
            )

            display_bnk["TYPE"] = (
                df_selisih_bnk["_BANK_TYPE"]
            )

            display_bnk["ISSUE"] = (
                df_selisih_bnk["STATUS_MATCH"]
            )

            st.dataframe(
                display_bnk,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue Bank. "
                "Semua transaksi bank memiliki pasangan FMSS."
            )

    # ========================================================
    # INVALID VA
    # ========================================================

    st.divider()

    with st.expander(
        "⚠️ Transaksi dengan VA Tidak Teridentifikasi",
        expanded=False
    ):

        iv1, iv2 = st.columns(2)

        # ----------------------------------------------------
        # INVALID FMSS
        # ----------------------------------------------------

        with iv1:

            st.markdown("### FMSS Invalid VA")

            if not df_invalid_int.empty:

                cols = [
                    col
                    for col in [
                        col_tanggal_int,
                        col_nominal_int,
                        col_keterangan_int
                    ]
                    if col in df_invalid_int.columns
                ]

                st.dataframe(
                    df_invalid_int[cols],
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.success(
                    "Tidak ada FMSS invalid VA."
                )

        # ----------------------------------------------------
        # INVALID BANK
        # ----------------------------------------------------

        with iv2:

            st.markdown("### Bank Invalid VA")

            if not df_invalid_bnk.empty:

                invalid_cols = []

                # Cari kolom tanggal/description yang tersedia
                for col in [
                    "_TANGGAL_DT",
                    "_BANK_TYPE",
                    "SOURCE_BANK"
                ]:
                    if col in df_invalid_bnk.columns:
                        invalid_cols.append(col)

                if "KODE_VA" in df_invalid_bnk.columns:
                    invalid_cols.append("KODE_VA")

                if "_CREDIT_NUM" in df_invalid_bnk.columns:
                    invalid_cols.append("_CREDIT_NUM")

                st.dataframe(
                    df_invalid_bnk[invalid_cols],
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.success(
                    "Tidak ada Bank invalid VA."
                )

    # ========================================================
    # DOWNLOAD REPORT
    # ========================================================

    st.divider()

    st.subheader("📥 Download Laporan")

    output = io.BytesIO()

    try:

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            # ------------------------------------------------
            # SUMMARY
            # ------------------------------------------------

            summary_export = pd.DataFrame({
                "METRIC": [
                    "Bank",
                    "Periode Rekonsiliasi",
                    "Matched",
                    "FMSS Only",
                    "Bank Only",
                    "FMSS Invalid VA",
                    "Bank Invalid VA",
                    "Match Rate FMSS",
                    "Match Rate Bank",
                    "Nominal Matched",
                    "Nominal FMSS Only",
                    "Nominal Bank Only"
                ],
                "VALUE": [
                    pilihan_bank,
                    safe_date_string(
                        st.session_state.recon_dates
                    ),
                    summary["matched_count"],
                    summary["fmss_only_count"],
                    summary["bank_only_count"],
                    summary["invalid_int_count"],
                    summary["invalid_bnk_count"],
                    f"{fmss_match_rate:.4f}%",
                    f"{bank_match_rate:.4f}%",
                    summary["matched_nominal"],
                    summary["fmss_only_nominal"],
                    summary["bank_only_nominal"]
                ]
            })

            summary_export.to_excel(
                writer,
                sheet_name="SUMMARY",
                index=False
            )

            # ------------------------------------------------
            # MATCHED
            # ------------------------------------------------

            if not df_matched.empty:

                export_matched = df_matched.copy()

                export_matched = export_matched.drop(
                    columns=[
                        "_STATUS_CLEAN",
                        "_TANGGAL_DT"
                    ],
                    errors="ignore"
                )

                export_matched.to_excel(
                    writer,
                    sheet_name="MATCHED_OK",
                    index=False
                )

            else:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada data matched."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="MATCHED_OK",
                    index=False
                )

            # ------------------------------------------------
            # FMSS ONLY
            # ------------------------------------------------

            if not df_selisih_int.empty:

                export_fmss = df_selisih_int.copy()

                export_fmss = export_fmss.drop(
                    columns=[
                        "_STATUS_CLEAN",
                        "_TANGGAL_DT"
                    ],
                    errors="ignore"
                )

                export_fmss.to_excel(
                    writer,
                    sheet_name="ISSUE_FMSS",
                    index=False
                )

            else:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada issue FMSS."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="ISSUE_FMSS",
                    index=False
                )

            # ------------------------------------------------
            # BANK ONLY
            # ------------------------------------------------

            if not df_selisih_bnk.empty:

                export_bank = df_selisih_bnk.copy()

                export_bank = export_bank.drop(
                    columns=[
                        "_TANGGAL_DT"
                    ],
                    errors="ignore"
                )

                export_bank.to_excel(
                    writer,
                    sheet_name="ISSUE_BANK",
                    index=False
                )

            else:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada issue Bank."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="ISSUE_BANK",
                    index=False
                )

            # ------------------------------------------------
            # INVALID FMSS
            # ------------------------------------------------

            if not df_invalid_int.empty:

                export_invalid_int = (
                    df_invalid_int.copy()
                )

                export_invalid_int.to_excel(
                    writer,
                    sheet_name="INVALID_FMSS",
                    index=False
                )

            else:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada FMSS invalid VA."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="INVALID_FMSS",
                    index=False
                )

            # ------------------------------------------------
            # INVALID BANK
            # ------------------------------------------------

            if not df_invalid_bnk.empty:

                export_invalid_bnk = (
                    df_invalid_bnk.copy()
                )

                export_invalid_bnk.to_excel(
                    writer,
                    sheet_name="INVALID_BANK",
                    index=False
                )

            else:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada Bank invalid VA."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="INVALID_BANK",
                    index=False
                )

        output.seek(0)

        st.download_button(
            label="📥 Download Laporan Lengkap (.xlsx)",
            data=output.getvalue(),
            file_name=(
                f"Laporan_Rekonsiliasi_"
                f"{pilihan_bank}_"
                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True
        )

    except Exception as e:

        st.error(
            "❌ Gagal membuat file Excel."
        )

        st.exception(e)


# ============================================================
# INFO JIKA BELUM LENGKAP
# ============================================================

elif pilihan_bank == "BRIVA":

    st.info(
        "💡 Upload **3 file** terlebih dahulu: "
        "FMSS, Mutasi BRIVA 57888, dan Mutasi BRIVA 57708."
    )

elif pilihan_bank != "":

    st.info(
        f"💡 Upload **2 file** terlebih dahulu: "
        f"FMSS dan Mutasi {pilihan_bank}."
    )

else:

    if file_int if "file_int" in locals() else False:
        st.info(
            "💡 Silakan pilih **Bank Sumber Mutasi** terlebih dahulu."
        )
