import streamlit as st
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIG
# =========================================================

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


# =========================================================
# SESSION STATE
# =========================================================

DEFAULT_STATE = {
    "sudah_diproses": False,
    "df_matched": pd.DataFrame(),
    "df_issue_fmss": pd.DataFrame(),
    "df_issue_bank": pd.DataFrame(),
    "df_invalid_fmss": pd.DataFrame(),
    "df_invalid_bank": pd.DataFrame(),
    "bank_terakhir": "",
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# HELPER
# =========================================================

def read_uploaded_file(uploaded_file):
    """
    Membaca CSV / XLSX secara fleksibel.
    """
    uploaded_file.seek(0)

    filename = uploaded_file.name.lower()

    if filename.endswith(".xlsx"):
        return pd.read_excel(uploaded_file)

    if filename.endswith(".csv"):
        # Coba autodetect separator terlebih dahulu
        return pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python"
        )

    raise ValueError(f"Format file tidak didukung: {uploaded_file.name}")


def normalize_text(value):
    """
    Normalisasi text agar lebih aman untuk pencocokan.
    """
    if pd.isna(value):
        return ""

    return str(value).strip()


def normalize_number(series):
    """
    Konversi nominal menjadi numeric secara aman.

    Mendukung:
    100000
    100.000
    100,000
    Rp 100.000
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    cleaned = (
        series.astype(str)
        .str.replace("Rp", "", regex=False)
        .str.replace("rp", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", "", regex=False)
    )

    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def extract_va(text):
    """
    Extract VA BRIVA Fastpay / Rajabiller.

    Prefix:
    57888 = Fastpay
    57708 = Rajabiller

    Mengambil 5 digit prefix + minimal 5 digit berikutnya.
    """
    if pd.isna(text):
        return None

    text = str(text)

    match = re.search(
        r"(57888\d{5,15}|57708\d{5,15})",
        text
    )

    if match:
        return match.group(1)

    return None


def classify_va(va):
    """
    Menentukan jenis BRIVA berdasarkan prefix.
    """
    if pd.isna(va) or va is None:
        return None

    va = str(va)

    if va.startswith("57888"):
        return "BRIVA FASTPAY"

    if va.startswith("57708"):
        return "BRIVA RAJABILLER"

    return None


def find_column(df, candidates):
    """
    Mencari nama kolom berdasarkan beberapa kandidat.
    """
    normalized = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    for candidate in candidates:
        candidate = candidate.lower()

        if candidate in normalized:
            return normalized[candidate]

    return None


def safe_sheet_name(name):
    """
    Membersihkan nama sheet Excel.
    """
    invalid_chars = r'[]:*?/\\'

    for char in invalid_chars:
        name = name.replace(char, "_")

    return name[:31]


def dataframe_to_excel(sheets):
    """
    Membuat XLSX menggunakan openpyxl langsung.
    Menghindari pandas ExcelWriter yang sebelumnya
    menyebabkan IndexError pada proses close().
    """
    wb = Workbook()

    # Hapus sheet default
    default_ws = wb.active
    wb.remove(default_ws)

    for sheet_name, df in sheets.items():

        ws = wb.create_sheet(
            safe_sheet_name(sheet_name)
        )

        if df is None or df.empty:
            ws.cell(row=1, column=1, value="Tidak ada data.")
            continue

        # Bersihkan nama kolom
        df_export = df.copy()

        columns = []
        for col in df_export.columns:
            col_name = str(col)

            # Hilangkan karakter bermasalah
            col_name = col_name.replace("\n", " ")
            col_name = col_name.replace("\r", " ")

            columns.append(col_name)

        df_export.columns = columns

        # Header
        for col_idx, col_name in enumerate(
            df_export.columns,
            start=1
        ):
            ws.cell(
                row=1,
                column=col_idx,
                value=col_name
            )

        # Data
        for row_idx, row in enumerate(
            df_export.itertuples(index=False, name=None),
            start=2
        ):
            for col_idx, value in enumerate(
                row,
                start=1
            ):

                if pd.isna(value):
                    value = None

                # Timestamp pandas → datetime biasa
                if isinstance(value, pd.Timestamp):
                    value = value.to_pydatetime()

                ws.cell(
                    row=row_idx,
                    column=col_idx,
                    value=value
                )

        # Freeze header
        ws.freeze_panes = "A2"

        # Auto width sederhana
        for col_idx, col_name in enumerate(
            df_export.columns,
            start=1
        ):
            max_length = len(str(col_name))

            # Batasi scan agar file besar tetap cepat
            sample = df_export.iloc[:500]

            for value in sample.iloc[:, col_idx - 1]:
                if pd.notna(value):
                    max_length = max(
                        max_length,
                        min(len(str(value)), 40)
                    )

            ws.column_dimensions[
                get_column_letter(col_idx)
            ].width = min(max_length + 2, 45)

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output.getvalue()


# =========================================================
# UPLOAD DATA
# =========================================================

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


# =========================================================
# RESET SAAT BANK BERUBAH
# =========================================================

if pilihan_bank != st.session_state.bank_terakhir:

    st.session_state.sudah_diproses = False

    st.session_state.df_matched = pd.DataFrame()
    st.session_state.df_issue_fmss = pd.DataFrame()
    st.session_state.df_issue_bank = pd.DataFrame()
    st.session_state.df_invalid_fmss = pd.DataFrame()
    st.session_state.df_invalid_bank = pd.DataFrame()

    st.session_state.bank_terakhir = pilihan_bank


# =========================================================
# UPLOAD
# =========================================================

st.subheader("2. Unggah File")

if pilihan_bank == "BRIVA":

    col1, col2, col3 = st.columns(3)

    with col1:
        file_int = st.file_uploader(
            "📄 FMSS",
            type=["csv", "xlsx"],
            key="fmss_briva"
        )

    with col2:
        file_bnk_fastpay = st.file_uploader(
            "🏦 BRIVA Fastpay — 57888",
            type=["csv", "xlsx"],
            key="briva_57888"
        )

    with col3:
        file_bnk_rajabiller = st.file_uploader(
            "🏦 BRIVA Rajabiller — 57708",
            type=["csv", "xlsx"],
            key="briva_57708"
        )

else:

    col1, col2 = st.columns(2)

    with col1:
        file_int = st.file_uploader(
            "📄 Data FMSS",
            type=["csv", "xlsx"],
            key="fmss_general"
        )

    with col2:
        file_bnk = st.file_uploader(
            f"🏦 Mutasi Bank {pilihan_bank}",
            type=["csv", "xlsx"],
            key="bank_general"
        )

    file_bnk_fastpay = None
    file_bnk_rajabiller = None


# =========================================================
# BRIVA CONFIG
# =========================================================

fee_fastpay = 1000
fee_rajabiller = 1000

if pilihan_bank == "BRIVA":

    st.subheader("3. Konfigurasi Fee")

    col1, col2 = st.columns(2)

    with col1:
        fee_fastpay = st.number_input(
            "Fee Fastpay (57888)",
            min_value=0,
            value=1000,
            step=100
        )

    with col2:
        fee_rajabiller = st.number_input(
            "Fee Rajabiller (57708)",
            min_value=0,
            value=1000,
            step=100
        )

    st.caption(
        "Rumus: Nominal FMSS + Fee = Nominal Mutasi Bank. "
        "Tanggal transaksi tidak dijadikan syarat utama matching "
        "karena data FMSS dan bank dapat berbeda tanggal."
    )


# =========================================================
# VALIDASI UPLOAD
# =========================================================

ready = False

if pilihan_bank == "BRIVA":

    ready = (
        file_int is not None
        and file_bnk_fastpay is not None
        and file_bnk_rajabiller is not None
    )

elif pilihan_bank != "":

    ready = (
        file_int is not None
        and file_bnk is not None
    )


# =========================================================
# PROSES
# =========================================================

if pilihan_bank != "":

    st.divider()

    if ready:

        if st.button(
            f"🚀 Mulai Croscek Data {pilihan_bank}",
            type="primary",
            use_container_width=True
        ):

            st.session_state.sudah_diproses = False

            start_message = st.empty()

            try:

                with st.spinner(
                    "Sedang memproses rekonsiliasi..."
                ):

                    # =================================================
                    # LOAD FMSS
                    # =================================================

                    df_int = read_uploaded_file(file_int)

                    df_int.columns = [
                        str(c).strip()
                        for c in df_int.columns
                    ]

                    # Cari kolom FMSS
                    status_col = find_column(
                        df_int,
                        [
                            "status"
                        ]
                    )

                    nominal_col = find_column(
                        df_int,
                        [
                            "nominal",
                            "amount",
                            "jumlah"
                        ]
                    )

                    keterangan_col = find_column(
                        df_int,
                        [
                            "keterangan",
                            "description",
                            "deskripsi"
                        ]
                    )

                    tanggal_int_col = find_column(
                        df_int,
                        [
                            "tanggal_transfer",
                            "tanggal",
                            "tgl_transfer",
                            "created_at"
                        ]
                    )

                    if status_col is None:
                        raise ValueError(
                            "Kolom 'status' tidak ditemukan pada FMSS."
                        )

                    if nominal_col is None:
                        raise ValueError(
                            "Kolom 'nominal' tidak ditemukan pada FMSS."
                        )

                    if keterangan_col is None:
                        raise ValueError(
                            "Kolom 'keterangan' tidak ditemukan pada FMSS."
                        )


                    # =================================================
                    # FILTER FMSS SUKSES
                    # =================================================

                    df_int = df_int[
                        df_int[status_col]
                        .astype(str)
                        .str.strip()
                        .str.upper()
                        .eq("SUKSES")
                    ].copy()

                    # Nominal numeric
                    df_int["NOMINAL_FMSS"] = normalize_number(
                        df_int[nominal_col]
                    )

                    # Extract VA
                    df_int["KODE_VA"] = (
                        df_int[keterangan_col]
                        .apply(extract_va)
                    )

                    # Jenis VA
                    df_int["JENIS_VA"] = (
                        df_int["KODE_VA"]
                        .apply(classify_va)
                    )

                    # =================================================
                    # INVALID VA FMSS
                    # =================================================

                    df_invalid_fmss = df_int[
                        df_int["KODE_VA"].isna()
                    ].copy()

                    # Hanya transaksi yang VA-nya dikenal
                    df_int_valid = df_int[
                        df_int["KODE_VA"].notna()
                    ].copy()


                    # =================================================
                    # BANK
                    # =================================================

                    if pilihan_bank == "BRIVA":

                        # ---------------------------------------------
                        # FASTPAY 57888
                        # ---------------------------------------------

                        df_fastpay = read_uploaded_file(
                            file_bnk_fastpay
                        )

                        df_fastpay.columns = [
                            str(c).strip()
                            for c in df_fastpay.columns
                        ]

                        # ---------------------------------------------
                        # RAJABILLER 57708
                        # ---------------------------------------------

                        df_rajabiller = read_uploaded_file(
                            file_bnk_rajabiller
                        )

                        df_rajabiller.columns = [
                            str(c).strip()
                            for c in df_rajabiller.columns
                        ]

                        # ---------------------------------------------
                        # VALIDATE BANK COLUMNS
                        # ---------------------------------------------

                        bank_kredit_fastpay = find_column(
                            df_fastpay,
                            [
                                "MUTASI_KREDIT",
                                "mutasi kredit",
                                "kredit",
                                "credit"
                            ]
                        )

                        bank_desc_fastpay = find_column(
                            df_fastpay,
                            [
                                "DESK_TRAN",
                                "desk tran",
                                "description",
                                "keterangan"
                            ]
                        )

                        bank_tanggal_fastpay = find_column(
                            df_fastpay,
                            [
                                "TANGGAL",
                                "tanggal",
                                "tgl",
                                "tanggal_transaksi",
                                "date"
                            ]
                        )

                        bank_kredit_rajabiller = find_column(
                            df_rajabiller,
                            [
                                "MUTASI_KREDIT",
                                "mutasi kredit",
                                "kredit",
                                "credit"
                            ]
                        )

                        bank_desc_rajabiller = find_column(
                            df_rajabiller,
                            [
                                "DESK_TRAN",
                                "desk tran",
                                "description",
                                "keterangan"
                            ]
                        )

                        bank_tanggal_rajabiller = find_column(
                            df_rajabiller,
                            [
                                "TANGGAL",
                                "tanggal",
                                "tgl",
                                "tanggal_transaksi",
                                "date"
                            ]
                        )

                        if bank_kredit_fastpay is None:
                            raise ValueError(
                                "Kolom MUTASI_KREDIT tidak ditemukan "
                                "pada file BRIVA 57888."
                            )

                        if bank_desc_fastpay is None:
                            raise ValueError(
                                "Kolom DESK_TRAN tidak ditemukan "
                                "pada file BRIVA 57888."
                            )

                        if bank_kredit_rajabiller is None:
                            raise ValueError(
                                "Kolom MUTASI_KREDIT tidak ditemukan "
                                "pada file BRIVA 57708."
                            )

                        if bank_desc_rajabiller is None:
                            raise ValueError(
                                "Kolom DESK_TRAN tidak ditemukan "
                                "pada file BRIVA 57708."
                            )


                        # =================================================
                        # PREPARE FASTPAY
                        # =================================================

                        df_fastpay["NOMINAL_BANK"] = normalize_number(
                            df_fastpay[bank_kredit_fastpay]
                        )

                        # Hanya uang masuk
                        df_fastpay = df_fastpay[
                            df_fastpay["NOMINAL_BANK"] > 0
                        ].copy()

                        df_fastpay["KODE_VA"] = (
                            df_fastpay[bank_desc_fastpay]
                            .apply(extract_va)
                        )

                        df_fastpay["JENIS_VA"] = "BRIVA FASTPAY"

                        df_fastpay["SUMBER_BANK"] = (
                            "BRIVA FASTPAY 57888"
                        )

                        if bank_tanggal_fastpay:
                            df_fastpay["TANGGAL_BANK"] = pd.to_datetime(
                                df_fastpay[bank_tanggal_fastpay],
                                errors="coerce"
                            )
                        else:
                            df_fastpay["TANGGAL_BANK"] = pd.NaT


                        # =================================================
                        # PREPARE RAJABILLER
                        # =================================================

                        df_rajabiller["NOMINAL_BANK"] = normalize_number(
                            df_rajabiller[bank_kredit_rajabiller]
                        )

                        df_rajabiller = df_rajabiller[
                            df_rajabiller["NOMINAL_BANK"] > 0
                        ].copy()

                        df_rajabiller["KODE_VA"] = (
                            df_rajabiller[bank_desc_rajabiller]
                            .apply(extract_va)
                        )

                        df_rajabiller["JENIS_VA"] = (
                            "BRIVA RAJABILLER"
                        )

                        df_rajabiller["SUMBER_BANK"] = (
                            "BRIVA RAJABILLER 57708"
                        )

                        if bank_tanggal_rajabiller:
                            df_rajabiller["TANGGAL_BANK"] = pd.to_datetime(
                                df_rajabiller[
                                    bank_tanggal_rajabiller
                                ],
                                errors="coerce"
                            )
                        else:
                            df_rajabiller["TANGGAL_BANK"] = pd.NaT


                        # =================================================
                        # INVALID VA BANK
                        # =================================================

                        invalid_fastpay = df_fastpay[
                            df_fastpay["KODE_VA"].isna()
                        ].copy()

                        invalid_rajabiller = df_rajabiller[
                            df_rajabiller["KODE_VA"].isna()
                        ].copy()

                        df_invalid_bank = pd.concat(
                            [
                                invalid_fastpay,
                                invalid_rajabiller
                            ],
                            ignore_index=True
                        )

                        # Bank valid
                        df_fastpay = df_fastpay[
                            df_fastpay["KODE_VA"].notna()
                        ].copy()

                        df_rajabiller = df_rajabiller[
                            df_rajabiller["KODE_VA"].notna()
                        ].copy()


                        # =================================================
                        # COMBINE BANK
                        # =================================================

                        df_bank = pd.concat(
                            [
                                df_fastpay,
                                df_rajabiller
                            ],
                            ignore_index=True
                        )

                    else:

                        # =================================================
                        # BANK NON BRIVA
                        # =================================================

                        df_bank = read_uploaded_file(
                            file_bnk
                        )

                        df_bank.columns = [
                            str(c).strip()
                            for c in df_bank.columns
                        ]

                        bank_kredit = find_column(
                            df_bank,
                            [
                                "MUTASI_KREDIT",
                                "mutasi kredit",
                                "kredit",
                                "credit"
                            ]
                        )

                        bank_desc = find_column(
                            df_bank,
                            [
                                "DESK_TRAN",
                                "desk tran",
                                "description",
                                "keterangan"
                            ]
                        )

                        if bank_kredit is None:
                            raise ValueError(
                                "Kolom kredit bank tidak ditemukan."
                            )

                        if bank_desc is None:
                            raise ValueError(
                                "Kolom deskripsi bank tidak ditemukan."
                            )

                        df_bank["NOMINAL_BANK"] = normalize_number(
                            df_bank[bank_kredit]
                        )

                        df_bank = df_bank[
                            df_bank["NOMINAL_BANK"] > 0
                        ].copy()

                        df_bank["KODE_VA"] = (
                            df_bank[bank_desc]
                            .apply(extract_va)
                        )

                        df_bank["JENIS_VA"] = pilihan_bank

                        df_bank["SUMBER_BANK"] = pilihan_bank

                        df_bank["TANGGAL_BANK"] = pd.NaT

                        df_invalid_bank = df_bank[
                            df_bank["KODE_VA"].isna()
                        ].copy()

                        df_bank = df_bank[
                            df_bank["KODE_VA"].notna()
                        ].copy()


                    # =================================================
                    # FEE
                    # =================================================

                    if pilihan_bank == "BRIVA":

                        df_int_valid["FEE"] = 0

                        df_int_valid.loc[
                            df_int_valid["JENIS_VA"]
                            == "BRIVA FASTPAY",
                            "FEE"
                        ] = fee_fastpay

                        df_int_valid.loc[
                            df_int_valid["JENIS_VA"]
                            == "BRIVA RAJABILLER",
                            "FEE"
                        ] = fee_rajabiller

                    else:

                        df_int_valid["FEE"] = 0


                    # =================================================
                    # NOMINAL MATCHING
                    # =================================================

                    df_int_valid["NOMINAL_MATCH"] = (
                        df_int_valid["NOMINAL_FMSS"]
                        + df_int_valid["FEE"]
                    )


                    # =================================================
                    # KEY MATCHING
                    #
                    # SANGAT PENTING:
                    # Tidak menggunakan tanggal sebagai key.
                    #
                    # Karena:
                    # FMSS = 17 Agustus
                    # Bank = 18 Agustus
                    #
                    # Tanggal hanya untuk diagnostic.
                    # =================================================

                    df_int_valid["_MATCH_KEY"] = (
                        df_int_valid["JENIS_VA"].fillna("")
                        .astype(str)
                        + "|"
                        + df_int_valid["KODE_VA"].fillna("")
                        .astype(str)
                        + "|"
                        + df_int_valid["NOMINAL_MATCH"]
                        .round(0)
                        .astype("int64")
                        .astype(str)
                    )

                    df_bank["_MATCH_KEY"] = (
                        df_bank["JENIS_VA"].fillna("")
                        .astype(str)
                        + "|"
                        + df_bank["KODE_VA"].fillna("")
                        .astype(str)
                        + "|"
                        + df_bank["NOMINAL_BANK"]
                        .round(0)
                        .astype("int64")
                        .astype(str)
                    )


                    # =================================================
                    # 1-TO-1 MATCHING
                    #
                    # cumcount() menjaga duplicate tetap 1-to-1.
                    #
                    # Contoh:
                    #
                    # FMSS:
                    # VA A | 1000
                    # VA A | 1000
                    #
                    # BANK:
                    # VA A | 2000
                    # VA A | 2000
                    #
                    # hasil:
                    # FMSS #1 -> BANK #1
                    # FMSS #2 -> BANK #2
                    # =================================================

                    df_int_valid["_MATCH_SEQ"] = (
                        df_int_valid
                        .groupby("_MATCH_KEY")
                        .cumcount()
                    )

                    df_bank["_MATCH_SEQ"] = (
                        df_bank
                        .groupby("_MATCH_KEY")
                        .cumcount()
                    )


                    # =================================================
                    # MERGE
                    # =================================================

                    bank_for_merge = df_bank.copy()

                    bank_for_merge = bank_for_merge.rename(
                        columns={
                            "NOMINAL_BANK": "MATCH_MUTASI_KREDIT",
                            "TANGGAL_BANK": "MATCH_TANGGAL_BANK",
                            "SUMBER_BANK": "MATCH_SUMBER_BANK"
                        }
                    )

                    bank_for_merge = bank_for_merge[
                        [
                            "_MATCH_KEY",
                            "_MATCH_SEQ",
                            "KODE_VA",
                            "JENIS_VA",
                            "MATCH_MUTASI_KREDIT",
                            "MATCH_TANGGAL_BANK",
                            "MATCH_SUMBER_BANK"
                        ]
                    ]

                    df_merged = df_int_valid.merge(
                        bank_for_merge,
                        on=[
                            "_MATCH_KEY",
                            "_MATCH_SEQ"
                        ],
                        how="left",
                        indicator=True,
                        suffixes=(
                            "",
                            "_BANK"
                        )
                    )


                    # =================================================
                    # MATCHED
                    # =================================================

                    df_matched = df_merged[
                        df_merged["_merge"] == "both"
                    ].copy()


                    # =================================================
                    # ISSUE FMSS
                    # =================================================

                    df_issue_fmss = df_merged[
                        df_merged["_merge"] == "left_only"
                    ].copy()

                    df_issue_fmss["ISSUE"] = "FMSS_ONLY"


                    # =================================================
                    # ISSUE BANK
                    #
                    # Bank yang tidak mendapatkan pasangan FMSS.
                    # =================================================

                    matched_keys = set(
                        df_matched["_MATCH_KEY"]
                        + "|"
                        + df_matched["_MATCH_SEQ"]
                        .astype(str)
                    )

                    df_bank["_MATCH_ID"] = (
                        df_bank["_MATCH_KEY"]
                        + "|"
                        + df_bank["_MATCH_SEQ"]
                        .astype(str)
                    )

                    df_issue_bank = df_bank[
                        ~df_bank["_MATCH_ID"].isin(
                            matched_keys
                        )
                    ].copy()

                    df_issue_bank["ISSUE"] = "BANK_ONLY"


                    # =================================================
                    # DATE DIAGNOSTIC
                    #
                    # Tidak digunakan untuk matching.
                    # =================================================

                    if tanggal_int_col:

                        df_matched["TANGGAL_FMSS"] = pd.to_datetime(
                            df_matched[tanggal_int_col],
                            errors="coerce"
                        )

                        df_matched["SELISIH_HARI"] = (
                            df_matched["MATCH_TANGGAL_BANK"]
                            .dt.normalize()
                            -
                            df_matched["TANGGAL_FMSS"]
                            .dt.normalize()
                        ).dt.days

                    else:

                        df_matched["TANGGAL_FMSS"] = pd.NaT
                        df_matched["SELISIH_HARI"] = pd.NA


                    # =================================================
                    # CLEAN MATCHED
                    # =================================================

                    df_matched["STATUS_MATCH"] = "MATCHED"

                    # Kolom penting ditaruh di depan
                    preferred_matched = [
                        "KODE_VA",
                        "JENIS_VA",
                        "NOMINAL_FMSS",
                        "FEE",
                        "NOMINAL_MATCH",
                        "MATCH_MUTASI_KREDIT",
                        "TANGGAL_FMSS",
                        "MATCH_TANGGAL_BANK",
                        "SELISIH_HARI",
                        "MATCH_SUMBER_BANK",
                        "STATUS_MATCH"
                    ]

                    existing = [
                        c for c in preferred_matched
                        if c in df_matched.columns
                    ]

                    remaining = [
                        c for c in df_matched.columns
                        if c not in existing
                        and not c.startswith("_")
                        and c not in ["_merge"]
                    ]

                    df_matched = df_matched[
                        existing + remaining
                    ]


                    # =================================================
                    # CLEAN ISSUE FMSS
                    # =================================================

                    preferred_issue_fmss = [
                        "KODE_VA",
                        "JENIS_VA",
                        "NOMINAL_FMSS",
                        "FEE",
                        "NOMINAL_MATCH",
                        "ISSUE"
                    ]

                    existing = [
                        c for c in preferred_issue_fmss
                        if c in df_issue_fmss.columns
                    ]

                    remaining = [
                        c for c in df_issue_fmss.columns
                        if c not in existing
                        and not c.startswith("_")
                        and c not in ["_merge"]
                    ]

                    df_issue_fmss = df_issue_fmss[
                        existing + remaining
                    ]


                    # =================================================
                    # CLEAN ISSUE BANK
                    # =================================================

                    preferred_issue_bank = [
                        "KODE_VA",
                        "JENIS_VA",
                        "NOMINAL_BANK",
                        "TANGGAL_BANK",
                        "SUMBER_BANK",
                        "ISSUE"
                    ]

                    existing = [
                        c for c in preferred_issue_bank
                        if c in df_issue_bank.columns
                    ]

                    remaining = [
                        c for c in df_issue_bank.columns
                        if c not in existing
                        and not c.startswith("_")
                    ]

                    df_issue_bank = df_issue_bank[
                        existing + remaining
                    ]


                    # =================================================
                    # SAVE STATE
                    # =================================================

                    st.session_state.df_matched = df_matched
                    st.session_state.df_issue_fmss = df_issue_fmss
                    st.session_state.df_issue_bank = df_issue_bank
                    st.session_state.df_invalid_fmss = df_invalid_fmss
                    st.session_state.df_invalid_bank = df_invalid_bank

                    st.session_state.sudah_diproses = True

            except Exception as e:

                st.session_state.sudah_diproses = False

                st.error(
                    f"❌ Rekonsiliasi gagal: {str(e)}"
                )

                st.exception(e)

    else:

        if pilihan_bank == "BRIVA":

            st.info(
                "💡 Upload 3 file terlebih dahulu: "
                "FMSS, Mutasi BRIVA 57888, dan Mutasi BRIVA 57708."
            )

        else:

            st.info(
                "💡 Upload data FMSS dan mutasi bank terlebih dahulu."
            )


# =========================================================
# HASIL
# =========================================================

if st.session_state.sudah_diproses:

    df_matched = st.session_state.df_matched
    df_issue_fmss = st.session_state.df_issue_fmss
    df_issue_bank = st.session_state.df_issue_bank
    df_invalid_fmss = st.session_state.df_invalid_fmss
    df_invalid_bank = st.session_state.df_invalid_bank


    # =====================================================
    # SUMMARY
    # =====================================================

    st.divider()

    st.subheader(
        f"🎯 Ringkasan Rekonsiliasi {pilihan_bank}"
    )

    m1, m2, m3, m4 = st.columns(4)

    m1.metric(
        "✅ Matched Sempurna",
        f"{len(df_matched):,} Trx"
    )

    m2.metric(
        "⚠️ Issue FMSS",
        f"{len(df_issue_fmss):,} Trx"
    )

    m3.metric(
        "⚠️ Issue Bank",
        f"{len(df_issue_bank):,} Trx"
    )

    m4.metric(
        "🚨 Invalid VA",
        f"{len(df_invalid_fmss) + len(df_invalid_bank):,} Trx"
    )


    # =====================================================
    # NOMINAL
    # =====================================================

    st.subheader("💰 Ringkasan Nominal")

    nominal_matched = (
        df_matched["NOMINAL_FMSS"].sum()
        if not df_matched.empty
        else 0
    )

    nominal_issue_fmss = (
        df_issue_fmss["NOMINAL_FMSS"].sum()
        if not df_issue_fmss.empty
        else 0
    )

    nominal_issue_bank = (
        df_issue_bank["NOMINAL_BANK"].sum()
        if not df_issue_bank.empty
        else 0
    )

    n1, n2, n3 = st.columns(3)

    n1.metric(
        "Matched",
        f"Rp {nominal_matched:,.0f}"
    )

    n2.metric(
        "Issue FMSS",
        f"Rp {nominal_issue_fmss:,.0f}"
    )

    n3.metric(
        "Issue Bank",
        f"Rp {nominal_issue_bank:,.0f}"
    )


    # =====================================================
    # DATE ANALYSIS
    # =====================================================

    if not df_matched.empty and "SELISIH_HARI" in df_matched.columns:

        date_diff = (
            df_matched["SELISIH_HARI"]
            .dropna()
        )

        if not date_diff.empty:

            count_plus_one = (
                date_diff == 1
            ).sum()

            count_same = (
                date_diff == 0
            ).sum()

            st.info(
                f"📅 Analisis tanggal: "
                f"{count_same:,} matched memiliki tanggal sama, "
                f"{count_plus_one:,} matched memiliki tanggal bank "
                f"+1 hari dari FMSS. "
                f"Perbedaan tanggal tidak digunakan sebagai syarat matching."
            )


    # =====================================================
    # ISSUE TABLE
    # =====================================================

    st.divider()

    col1, col2 = st.columns(2)


    # -----------------------------------------------------
    # ISSUE FMSS
    # -----------------------------------------------------

    with col1:

        st.subheader("🚨 Issue FMSS")

        if df_issue_fmss.empty:

            st.success(
                "Tidak ada issue FMSS. "
                "Semua transaksi FMSS memiliki pasangan."
            )

        else:

            display_cols = [
                c for c in [
                    "KODE_VA",
                    "JENIS_VA",
                    "NOMINAL_FMSS",
                    "FEE",
                    "NOMINAL_MATCH",
                    "ISSUE"
                ]
                if c in df_issue_fmss.columns
            ]

            st.dataframe(
                df_issue_fmss[
                    display_cols
                ].head(500),
                use_container_width=True,
                hide_index=True
            )

            if len(df_issue_fmss) > 500:
                st.caption(
                    f"Menampilkan 500 dari "
                    f"{len(df_issue_fmss):,} transaksi. "
                    f"Data lengkap tersedia di Excel."
                )


    # -----------------------------------------------------
    # ISSUE BANK
    # -----------------------------------------------------

    with col2:

        st.subheader("🚨 Issue Bank")

        if df_issue_bank.empty:

            st.success(
                "Tidak ada issue Bank. "
                "Semua transaksi bank memiliki pasangan."
            )

        else:

            display_cols = [
                c for c in [
                    "KODE_VA",
                    "JENIS_VA",
                    "NOMINAL_BANK",
                    "TANGGAL_BANK",
                    "SUMBER_BANK",
                    "ISSUE"
                ]
                if c in df_issue_bank.columns
            ]

            st.dataframe(
                df_issue_bank[
                    display_cols
                ].head(500),
                use_container_width=True,
                hide_index=True
            )

            if len(df_issue_bank) > 500:
                st.caption(
                    f"Menampilkan 500 dari "
                    f"{len(df_issue_bank):,} transaksi. "
                    f"Data lengkap tersedia di Excel."
                )


    # =====================================================
    # INVALID VA
    # =====================================================

    st.divider()

    with st.expander(
        "⚠️ Transaksi dengan VA Tidak Teridentifikasi"
    ):

        c1, c2 = st.columns(2)

        with c1:

            st.markdown("### FMSS Invalid VA")

            if df_invalid_fmss.empty:

                st.success(
                    "Tidak ada FMSS invalid VA."
                )

            else:

                st.dataframe(
                    df_invalid_fmss.head(500),
                    use_container_width=True,
                    hide_index=True
                )

        with c2:

            st.markdown("### Bank Invalid VA")

            if df_invalid_bank.empty:

                st.success(
                    "Tidak ada Bank invalid VA."
                )

            else:

                st.dataframe(
                    df_invalid_bank.head(500),
                    use_container_width=True,
                    hide_index=True
                )


    # =====================================================
    # DOWNLOAD
    # =====================================================

    st.divider()

    st.subheader("📥 Download Laporan")

    excel_data = dataframe_to_excel(
        {
            "MATCHED_OK": df_matched,
            "ISSUE_FMSS": df_issue_fmss,
            "ISSUE_BANK": df_issue_bank,
            "INVALID_FMSS": df_invalid_fmss,
            "INVALID_BANK": df_invalid_bank,
        }
    )

    st.download_button(
        label="📥 Download Laporan Lengkap (.xlsx)",
        data=excel_data,
        file_name=(
            f"Laporan_Rekonsiliasi_{pilihan_bank}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True
    )
