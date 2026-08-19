import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime, timedelta
from collections import defaultdict, deque

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
    "Dashboard rekonsiliasi otomatis antara data FMSS "
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
    "df_out_period": pd.DataFrame(),
    "recon_dates": [],
    "summary": {},
    "pilihan_bank_terakhir": ""
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# HELPER
# ============================================================

def find_column(df, candidates, required=True):
    if df is None or df.empty:
        if required:
            raise ValueError(
                f"Data kosong. Tidak dapat mencari kolom: {candidates}"
            )
        return None

    for candidate in candidates:
        if candidate in df.columns:
            return candidate

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


@st.cache_data(show_spinner=False)
def read_uploaded_file_cached(file_bytes, filename):
    filename_lower = filename.lower()

    if filename_lower.endswith(".csv"):
        # Coba parser C terlebih dahulu agar file besar jauh lebih cepat.
        # Fallback ke python engine hanya jika struktur CSV tidak standar.
        try:
            return pd.read_csv(
                io.BytesIO(file_bytes),
                low_memory=False
            )
        except Exception:
            # Python engine TIDAK mendukung low_memory.
            # Jangan kirim parameter tersebut ke fallback.
            return pd.read_csv(
                io.BytesIO(file_bytes),
                sep=None,
                engine="python"
            )

    if filename_lower.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(file_bytes))

    raise ValueError(
        f"Format file tidak didukung: {filename}"
    )


def read_uploaded_file(uploaded_file):
    uploaded_file.seek(0)
    data = uploaded_file.getvalue()
    return read_uploaded_file_cached(data, uploaded_file.name)


def clean_numeric(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(
            series,
            errors="coerce"
        ).fillna(0)

    s = (
        series.astype("string")
        .str.strip()
        .str.replace("Rp", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    # Format umum Indonesia:
    # 1.234.567,89 -> 1234567.89
    # 1234567 -> 1234567
    # 123,456 -> 123456
    #
    # Untuk data bank/FMMSS yang umumnya integer:
    # hapus separator ribuan.
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(".", "", regex=False)

    return pd.to_numeric(
        s,
        errors="coerce"
    ).fillna(0)


def parse_datetime(series):
    return pd.to_datetime(
        series,
        errors="coerce"
    )


# ============================================================
# VA PARSER
# ============================================================

def extract_va(text):
    """
    Parser VA umum untuk BRIVA.

    BRIVA:
      57888 + 5-15 digit
      57708 + 5-15 digit

    BNIVA:
      ditangani khusus oleh extract_bniva_va()
    """
    if pd.isna(text):
        return None

    text = str(text)

    match = re.search(
        r"(57(?:888|708)\d{5,15})",
        text
    )

    if match:
        return match.group(1)

    return None


def extract_bniva_va(text):
    """
    STRICT BNIVA VA.

    Berdasarkan cross-check file BNIVA aktual:
        988765 + 10 digit
        total = 16 digit

    Jangan mengambil digit tambahan setelah VA.
    """
    if pd.isna(text):
        return None

    text = str(text)

    match = re.search(
        r"(?<!\d)(988765\d{10})(?!\d)",
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

    if va.startswith("988765"):
        return "BNIVA"

    return "UNKNOWN"


def classify_bank_transaction(description):
    text = str(description).upper()

    if "ATM" in text:
        return "ATM / MANUAL"

    if "TRF BERSAMA" in text:
        return "TRANSFER / MANUAL"

    if "BRIVA" in text:
        return "BRIVA"

    if "BFVA" in text:
        return "BFVA"

    if "BNI" in text or "BNIVA" in text:
        return "BNIVA"

    if "VA" in text:
        return "VA"

    return "OTHER"


def format_rupiah(value):
    try:
        value = float(value)
    except Exception:
        value = 0

    return "Rp {:,.0f}".format(value).replace(",", ".")


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


def classify_issue_bank(description):
    category = classify_bank_transaction(description)

    if category in ["ATM / MANUAL", "TRANSFER / MANUAL"]:
        return "BANK_ONLY - MANUAL/ATM"

    if category == "BRIVA":
        return "BANK_ONLY - BRIVA"

    if category == "BFVA":
        return "BANK_ONLY - BFVA"

    if category == "BNIVA":
        return "BANK_ONLY - BNIVA"

    return "BANK_ONLY - OTHER"


def make_match_key(va, nominal):
    """
    Key utama rekonsiliasi:
        KODE_VA + NOMINAL

    Tanggal sengaja TIDAK dimasukkan karena:
    - FMSS 18 Aug dapat muncul di bank 17/18/19 Aug
    - posting/cutoff bank berbeda
    """
    return (
        str(va).strip(),
        int(round(float(nominal)))
    )


# ============================================================
# UI - BANK
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

if pilihan_bank != st.session_state.pilihan_bank_terakhir:
    st.session_state.sudah_diproses = False
    st.session_state.df_matched = pd.DataFrame()
    st.session_state.df_selisih_int = pd.DataFrame()
    st.session_state.df_selisih_bnk = pd.DataFrame()
    st.session_state.df_invalid_int = pd.DataFrame()
    st.session_state.df_invalid_bnk = pd.DataFrame()
    st.session_state.df_out_period = pd.DataFrame()
    st.session_state.recon_dates = []
    st.session_state.summary = {}
    st.session_state.pilihan_bank_terakhir = pilihan_bank


# ============================================================
# UPLOAD
# ============================================================

st.subheader("2. Unggah File")

file_int = None
file_bnk_general = None
file_bnk_57888 = None
file_bnk_57708 = None

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

elif pilihan_bank == "BNIVA":

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 📄 FMSS")
        file_int = st.file_uploader(
            "Upload data FMSS BNIVA",
            type=["csv", "xlsx"],
            key="fmss_bniva"
        )

    with col2:
        st.markdown("### 🏦 BNIVA")
        file_bnk_general = st.file_uploader(
            "Upload mutasi BNIVA",
            type=["csv", "xlsx"],
            key="bniva_general"
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


# ============================================================
# FEE
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
        "Rumus BRIVA: Nominal FMSS + Fee = Nominal mutasi bank."
    )

elif pilihan_bank == "BNIVA":

    st.subheader("3. Konfigurasi Fee")

    st.number_input(
        "Fee BNIVA",
        min_value=0,
        value=0,
        step=100,
        format="%d",
        disabled=True
    )

    st.caption(
        "BNIVA: berdasarkan cross-check data aktual, "
        "matching menggunakan Nominal FMSS = Nominal bank (Fee Rp0)."
    )

    fee_57888 = 1000
    fee_57708 = 1000

else:

    fee_57888 = 0
    fee_57708 = 0


# ============================================================
# PROCESS BUTTON
# ============================================================

can_process = False

if pilihan_bank == "BRIVA":
    can_process = bool(
        file_int and
        file_bnk_57888 and
        file_bnk_57708
    )

elif pilihan_bank in ["BNIVA", "BCAVA", "MANDIRIVA", "BSIVA", "MuamalatVA"]:
    can_process = bool(
        file_int and
        file_bnk_general
    )


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

                df_int = df_int.copy()

                df_int["_STATUS_CLEAN"] = (
                    df_int[col_status]
                    .astype("string")
                    .str.strip()
                    .str.upper()
                )

                df_int_sukses = df_int[
                    df_int["_STATUS_CLEAN"] == "SUKSES"
                ].copy()

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
                # EXTRACT FMSS VA
                # ====================================================

                if pilihan_bank == "BNIVA":
                    df_int_sukses["KODE_VA"] = (
                        df_int_sukses[col_keterangan_int]
                        .apply(extract_bniva_va)
                    )
                else:
                    df_int_sukses["KODE_VA"] = (
                        df_int_sukses[col_keterangan_int]
                        .apply(extract_va)
                    )

                df_int_sukses["JENIS_VA"] = (
                    df_int_sukses["KODE_VA"]
                    .apply(classify_va)
                )

                df_invalid_int = df_int_sukses[
                    df_int_sukses["KODE_VA"].isna()
                ].copy()

                df_int_valid = df_int_sukses[
                    df_int_sukses["KODE_VA"].notna()
                ].copy()

                df_int_valid["NOMINAL_ASLI"] = clean_numeric(
                    df_int_valid[col_nominal_int]
                )

                # ====================================================
                # EXPECTED BANK
                # ====================================================

                df_int_valid["EXPECTED_BANK"] = (
                    df_int_valid["NOMINAL_ASLI"]
                )

                if pilihan_bank == "BRIVA":

                    mask_57888 = (
                        df_int_valid["JENIS_VA"]
                        == "BRIVA FASTPAY"
                    )

                    mask_57708 = (
                        df_int_valid["JENIS_VA"]
                        == "BRIVA RAJABILLER"
                    )

                    df_int_valid.loc[
                        mask_57888,
                        "EXPECTED_BANK"
                    ] = (
                        df_int_valid.loc[
                            mask_57888,
                            "NOMINAL_ASLI"
                        ] + fee_57888
                    )

                    df_int_valid.loc[
                        mask_57708,
                        "EXPECTED_BANK"
                    ] = (
                        df_int_valid.loc[
                            mask_57708,
                            "NOMINAL_ASLI"
                        ] + fee_57708
                    )

                # ====================================================
                # BANK LOADING
                # ====================================================

                bank_sources = []

                def prepare_bank_file(
                    uploaded_file,
                    source_bank,
                    va_parser
                ):
                    df = read_uploaded_file(uploaded_file).copy()

                    col_credit = find_column(
                        df,
                        [
                            "MUTASI_KREDIT",
                            "mutasi_kredit",
                            "KREDIT",
                            "kredit"
                        ]
                    )

                    col_desc = find_column(
                        df,
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
                        df,
                        [
                            "TGL_TRAN",
                            "tgl_tran",
                            "TANGGAL_TRAN",
                            "tanggal_tran",
                            "TANGGAL",
                            "tanggal",
                            "POSTING_DATE",
                            "posting_date"
                        ]
                    )

                    df["_BANK_TYPE"] = (
                        df[col_desc]
                        .apply(classify_bank_transaction)
                    )

                    df["_TANGGAL_DT"] = parse_datetime(
                        df[col_date]
                    )

                    df["_CREDIT_NUM"] = clean_numeric(
                        df[col_credit]
                    )

                    # IMPORTANT:
                    # Jangan membatasi bank hanya ke tanggal FMSS.
                    # Posting bank bisa H-1 / H / H+1.
                    # Kita load seluruh credit valid, lalu tanggal
                    # hanya dipakai untuk audit OUT_OF_PERIOD.
                    df = df[
                        df["_CREDIT_NUM"] > 0
                    ].copy()

                    df["KODE_VA"] = (
                        df[col_desc]
                        .apply(va_parser)
                    )

                    df["JENIS_VA"] = (
                        df["KODE_VA"]
                        .apply(classify_va)
                    )

                    df["SOURCE_BANK"] = source_bank
                    df["_DESC_COLUMN"] = col_desc
                    df["_DATE_COLUMN"] = col_date

                    return df

                if pilihan_bank == "BRIVA":

                    bank_sources.append(
                        prepare_bank_file(
                            file_bnk_57888,
                            "BRIVA FASTPAY 57888",
                            extract_va
                        )
                    )

                    bank_sources.append(
                        prepare_bank_file(
                            file_bnk_57708,
                            "BRIVA RAJABILLER 57708",
                            extract_va
                        )
                    )

                elif pilihan_bank == "BNIVA":

                    bank_sources.append(
                        prepare_bank_file(
                            file_bnk_general,
                            "BNIVA",
                            extract_bniva_va
                        )
                    )

                else:

                    bank_sources.append(
                        prepare_bank_file(
                            file_bnk_general,
                            pilihan_bank,
                            extract_va
                        )
                    )

                df_bank = pd.concat(
                    bank_sources,
                    ignore_index=True
                )

                # ====================================================
                # INVALID BANK
                # ====================================================

                df_invalid_bnk = df_bank[
                    df_bank["KODE_VA"].isna()
                ].copy()

                df_bank_valid = df_bank[
                    df_bank["KODE_VA"].notna()
                ].copy()

                # ====================================================
                # FAST 1-TO-1 MATCHING
                # ====================================================
                #
                # OLD:
                # for each FMSS -> scan every bank row
                # O(N*M)
                #
                # NEW:
                # build dictionary:
                # (VA, nominal) -> deque(bank indexes)
                #
                # O(N+M)
                #
                # This is the main performance improvement.

                bank_queues = defaultdict(deque)

                for idx, row in df_bank_valid.iterrows():

                    key = make_match_key(
                        row["KODE_VA"],
                        row["_CREDIT_NUM"]
                    )

                    bank_queues[key].append(idx)

                matched_records = []
                unmatched_internal = []

                bank_used = set()

                for _, int_row in df_int_valid.iterrows():

                    key = make_match_key(
                        int_row["KODE_VA"],
                        int_row["EXPECTED_BANK"]
                    )

                    queue = bank_queues.get(key)

                    if queue:

                        bank_idx = queue.popleft()

                        if not queue:
                            bank_queues.pop(key, None)

                        bank_used.add(bank_idx)

                        bank_row = df_bank_valid.loc[
                            bank_idx
                        ]

                        record = int_row.to_dict()

                        record["MATCH_MUTASI_KREDIT"] = (
                            bank_row["_CREDIT_NUM"]
                        )

                        record["MATCH_DESK_TRAN"] = (
                            bank_row[
                                bank_row["_DESC_COLUMN"]
                            ]
                        )

                        record["MATCH_TANGGAL_BANK"] = (
                            bank_row["_TANGGAL_DT"]
                        )

                        record["SOURCE_BANK"] = (
                            bank_row["SOURCE_BANK"]
                        )

                        record["BANK_TYPE"] = (
                            bank_row["_BANK_TYPE"]
                        )

                        record["STATUS_MATCH"] = "MATCHED"

                        # Audit tanggal:
                        # tanggal bank boleh H-1/H/H+1.
                        bank_date = bank_row["_TANGGAL_DT"]

                        if pd.notna(bank_date):
                            bank_date_only = bank_date.date()
                            fmss_date = (
                                int_row["_TANGGAL_DT"].date()
                            )

                            delta_days = (
                                bank_date_only - fmss_date
                            ).days

                            record["SELISIH_HARI_BANK"] = (
                                delta_days
                            )

                            if delta_days == 0:
                                record["MATCH_DATE_STATUS"] = "SAME_DATE"

                            elif delta_days == -1:
                                record["MATCH_DATE_STATUS"] = "H-1"

                            elif delta_days == 1:
                                record["MATCH_DATE_STATUS"] = "H+1"

                            else:
                                record["MATCH_DATE_STATUS"] = (
                                    "OUT_OF_PERIOD"
                                )

                        else:
                            record["SELISIH_HARI_BANK"] = None
                            record["MATCH_DATE_STATUS"] = (
                                "DATE_UNKNOWN"
                            )

                        matched_records.append(record)

                    else:

                        record = int_row.to_dict()
                        record["STATUS_MATCH"] = "FMSS_ONLY"

                        unmatched_internal.append(record)

                # ====================================================
                # BANK REMAINING
                # ====================================================

                unmatched_bank = []
                out_period_bank = []

                recon_date_set = set(recon_dates)

                for idx, bank_row in df_bank_valid.iterrows():

                    if idx in bank_used:
                        continue

                    record = bank_row.to_dict()

                    bank_date = bank_row["_TANGGAL_DT"]

                    if pd.notna(bank_date):

                        bank_date_only = bank_date.date()

                        # Audit:
                        # di luar tanggal FMSS
                        if bank_date_only not in recon_date_set:

                            record["STATUS_MATCH"] = (
                                "OUT_OF_PERIOD"
                            )

                            out_period_bank.append(record)

                            continue

                    record["STATUS_MATCH"] = (
                        classify_issue_bank(
                            bank_row["_BANK_TYPE"]
                        )
                    )

                    unmatched_bank.append(record)

                # ====================================================
                # DATAFRAME RESULT
                # ====================================================

                df_matched = pd.DataFrame(
                    matched_records
                )

                df_selisih_int = pd.DataFrame(
                    unmatched_internal
                )

                df_selisih_bnk = pd.DataFrame(
                    unmatched_bank
                )

                df_out_period = pd.DataFrame(
                    out_period_bank
                )

                # ====================================================
                # SUMMARY
                # ====================================================

                matched_count = len(df_matched)
                fmss_only_count = len(df_selisih_int)
                bank_only_count = len(df_selisih_bnk)
                out_period_count = len(df_out_period)

                invalid_int_count = len(df_invalid_int)
                invalid_bnk_count = len(df_invalid_bnk)

                matched_nominal = (
                    df_matched["NOMINAL_ASLI"].sum()
                    if (
                        not df_matched.empty
                        and "NOMINAL_ASLI" in df_matched.columns
                    )
                    else 0
                )

                fmss_only_nominal = (
                    df_selisih_int["NOMINAL_ASLI"].sum()
                    if (
                        not df_selisih_int.empty
                        and "NOMINAL_ASLI" in df_selisih_int.columns
                    )
                    else 0
                )

                bank_only_nominal = (
                    df_selisih_bnk["_CREDIT_NUM"].sum()
                    if (
                        not df_selisih_bnk.empty
                        and "_CREDIT_NUM" in df_selisih_bnk.columns
                    )
                    else 0
                )

                out_period_nominal = (
                    df_out_period["_CREDIT_NUM"].sum()
                    if (
                        not df_out_period.empty
                        and "_CREDIT_NUM" in df_out_period.columns
                    )
                    else 0
                )

                summary = {
                    "matched_count": matched_count,
                    "fmss_only_count": fmss_only_count,
                    "bank_only_count": bank_only_count,
                    "out_period_count": out_period_count,
                    "invalid_int_count": invalid_int_count,
                    "invalid_bnk_count": invalid_bnk_count,
                    "matched_nominal": matched_nominal,
                    "fmss_only_nominal": fmss_only_nominal,
                    "bank_only_nominal": bank_only_nominal,
                    "out_period_nominal": out_period_nominal
                }

                # ====================================================
                # SAVE STATE
                # ====================================================

                st.session_state.df_matched = df_matched
                st.session_state.df_selisih_int = df_selisih_int
                st.session_state.df_selisih_bnk = df_selisih_bnk
                st.session_state.df_invalid_int = df_invalid_int
                st.session_state.df_invalid_bnk = df_invalid_bnk
                st.session_state.df_out_period = df_out_period
                st.session_state.summary = summary
                st.session_state.sudah_diproses = True

        except Exception as e:

            st.session_state.sudah_diproses = False

            st.error(
                "❌ Terjadi kesalahan saat memproses data."
            )

            st.exception(e)


# ============================================================
# RESULT
# ============================================================

if st.session_state.sudah_diproses:

    df_matched = st.session_state.df_matched
    df_selisih_int = st.session_state.df_selisih_int
    df_selisih_bnk = st.session_state.df_selisih_bnk
    df_invalid_int = st.session_state.df_invalid_int
    df_invalid_bnk = st.session_state.df_invalid_bnk
    df_out_period = st.session_state.df_out_period
    summary = st.session_state.summary

    st.divider()

    st.subheader(
        f"🎯 Ringkasan Rekonsiliasi {pilihan_bank}"
    )

    st.caption(
        f"Periode FMSS: "
        f"**{safe_date_string(st.session_state.recon_dates)}**"
    )

    # ========================================================
    # METRICS
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
        "📅 Di Luar Periode",
        f"{summary['out_period_count']:,} Trx"
    )

    # ========================================================
    # MATCH RATE
    # ========================================================

    total_fmss_valid = (
        len(df_matched) +
        len(df_selisih_int)
    )

    total_bank_valid_in_scope = (
        len(df_matched) +
        len(df_selisih_bnk)
    )

    fmss_match_rate = (
        len(df_matched) /
        total_fmss_valid *
        100
        if total_fmss_valid > 0
        else 0
    )

    bank_match_rate = (
        len(df_matched) /
        total_bank_valid_in_scope *
        100
        if total_bank_valid_in_scope > 0
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
    # NOMINAL
    # ========================================================

    st.subheader("💰 Ringkasan Nominal")

    n1, n2, n3, n4 = st.columns(4)

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

    n4.metric(
        "Di Luar Periode",
        format_rupiah(
            summary["out_period_nominal"]
        )
    )

    # ========================================================
    # BRIVA / BNIVA VALIDATION NOTE
    # ========================================================

    if pilihan_bank == "BNIVA":

        if (
            summary["matched_count"] == 806
            and summary["fmss_only_count"] == 0
        ):
            st.success(
                "✅ Validasi BNIVA: 806/806 transaksi FMSS matched."
            )

        else:
            st.warning(
                "⚠️ Hasil BNIVA berbeda dari baseline 806/806. "
                "Periksa detail issue sebelum digunakan."
            )

    # ========================================================
    # ISSUE TABLES
    # ========================================================

    st.divider()

    col_issue1, col_issue2 = st.columns(2)

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
                "Tidak ada issue FMSS."
            )

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

            display_bnk["TANGGAL BANK"] = (
                df_selisih_bnk["_TANGGAL_DT"]
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
                "Tidak ada issue Bank dalam periode."
            )

    # ========================================================
    # OUT OF PERIOD
    # ========================================================

    st.divider()

    with st.expander(
        "📅 Transaksi Bank Di Luar Periode FMSS",
        expanded=False
    ):

        if not df_out_period.empty:

            display_out = pd.DataFrame()

            display_out["KODE VA"] = (
                df_out_period["KODE_VA"]
            )

            display_out["JENIS VA"] = (
                df_out_period["JENIS_VA"]
            )

            display_out["NOMINAL BANK"] = (
                df_out_period["_CREDIT_NUM"]
            )

            display_out["TANGGAL BANK"] = (
                df_out_period["_TANGGAL_DT"]
            )

            display_out["SOURCE BANK"] = (
                df_out_period["SOURCE_BANK"]
            )

            display_out["ISSUE"] = "OUT_OF_PERIOD"

            st.dataframe(
                display_out,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada transaksi bank di luar periode."
            )

    # ========================================================
    # INVALID VA
    # ========================================================

    with st.expander(
        "⚠️ Transaksi dengan VA Tidak Teridentifikasi",
        expanded=False
    ):

        iv1, iv2 = st.columns(2)

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

        with iv2:

            st.markdown("### Bank Invalid VA")

            if not df_invalid_bnk.empty:

                invalid_cols = [
                    col
                    for col in [
                        "_TANGGAL_DT",
                        "_BANK_TYPE",
                        "SOURCE_BANK",
                        "KODE_VA",
                        "_CREDIT_NUM"
                    ]
                    if col in df_invalid_bnk.columns
                ]

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
    # DOWNLOAD
    # ========================================================

    st.divider()

    st.subheader("📥 Download Laporan")

    output = io.BytesIO()

    try:

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            summary_export = pd.DataFrame({
                "METRIC": [
                    "Bank",
                    "Periode FMSS",
                    "Matched",
                    "FMSS Only",
                    "Bank Only",
                    "Out of Period",
                    "FMSS Invalid VA",
                    "Bank Invalid VA",
                    "Match Rate FMSS",
                    "Match Rate Bank",
                    "Nominal Matched",
                    "Nominal FMSS Only",
                    "Nominal Bank Only",
                    "Nominal Out of Period"
                ],
                "VALUE": [
                    pilihan_bank,
                    safe_date_string(
                        st.session_state.recon_dates
                    ),
                    summary["matched_count"],
                    summary["fmss_only_count"],
                    summary["bank_only_count"],
                    summary["out_period_count"],
                    summary["invalid_int_count"],
                    summary["invalid_bnk_count"],
                    f"{fmss_match_rate:.4f}%",
                    f"{bank_match_rate:.4f}%",
                    summary["matched_nominal"],
                    summary["fmss_only_nominal"],
                    summary["bank_only_nominal"],
                    summary["out_period_nominal"]
                ]
            })

            summary_export.to_excel(
                writer,
                sheet_name="SUMMARY",
                index=False
            )

            def write_sheet(df, sheet_name, empty_message):
                if not df.empty:
                    export_df = df.copy()

                    export_df = export_df.drop(
                        columns=[
                            "_STATUS_CLEAN",
                            "_TANGGAL_DT",
                            "_DESC_COLUMN",
                            "_DATE_COLUMN"
                        ],
                        errors="ignore"
                    )

                    export_df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False
                    )
                else:
                    pd.DataFrame({
                        "INFO": [empty_message]
                    }).to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False
                    )

            write_sheet(
                df_matched,
                "MATCHED_OK",
                "Tidak ada data matched."
            )

            write_sheet(
                df_selisih_int,
                "ISSUE_FMSS",
                "Tidak ada issue FMSS."
            )

            write_sheet(
                df_selisih_bnk,
                "ISSUE_BANK",
                "Tidak ada issue Bank."
            )

            write_sheet(
                df_out_period,
                "OUT_OF_PERIOD",
                "Tidak ada transaksi bank di luar periode."
            )

            write_sheet(
                df_invalid_int,
                "INVALID_FMSS",
                "Tidak ada FMSS invalid VA."
            )

            write_sheet(
                df_invalid_bnk,
                "INVALID_BANK",
                "Tidak ada Bank invalid VA."
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
# INFO
# ============================================================

elif pilihan_bank == "BRIVA":

    st.info(
        "💡 Upload 3 file: FMSS, Mutasi BRIVA 57888, "
        "dan Mutasi BRIVA 57708."
    )

elif pilihan_bank in [
    "BNIVA",
    "BCAVA",
    "MANDIRIVA",
    "BSIVA",
    "MuamalatVA"
]:

    st.info(
        f"💡 Upload 2 file: FMSS dan Mutasi {pilihan_bank}."
    )

else:

    st.info(
        "💡 Silakan pilih Bank Sumber Mutasi terlebih dahulu."
    )
