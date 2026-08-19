import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
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
    "Dashboard rekonsiliasi otomatis antara data deposit FMSS dengan mutasi bank."
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
    "df_out_of_period": pd.DataFrame(),
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
# HELPER
# ============================================================

def find_column(df, candidates, required=True):
    if df is None or df.empty:
        if required:
            raise ValueError(f"Data kosong. Tidak dapat mencari kolom: {candidates}")
        return None

    for candidate in candidates:
        if candidate in df.columns:
            return candidate

    mapping = {str(c).strip().lower(): c for c in df.columns}

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
    CSV tetap memakai Python engine karena file export FMSS tertentu
    dapat memiliki baris sangat panjang / format yang membuat C engine gagal.
    XLSX tetap memakai pandas Excel reader.
    """
    if uploaded_file is None:
        raise ValueError("File belum dipilih.")

    uploaded_file.seek(0)
    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        return pd.read_csv(
            uploaded_file,
            sep=",",
            engine="python",
            skip_blank_lines=True
        )

    if filename.endswith(".xlsx"):
        return pd.read_excel(uploaded_file)

    raise ValueError(f"Format file tidak didukung: {uploaded_file.name}")


def clean_numeric(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    cleaned = (
        series.astype("string")
        .str.replace("Rp", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
    )

    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def parse_datetime(series, fmt=None):
    if fmt:
        result = pd.to_datetime(series, format=fmt, errors="coerce")
        if result.notna().any():
            return result

    return pd.to_datetime(series, errors="coerce")


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
    except Exception:
        value = 0
    return "Rp {:,.0f}".format(value).replace(",", ".")


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
    if "VA" in text:
        return "VA"
    return "OTHER"


def classify_issue_bank(description):
    category = classify_bank_transaction(description)

    if category in ["ATM / MANUAL", "TRANSFER / MANUAL"]:
        return "BANK_ONLY - MANUAL/ATM"
    if category == "BRIVA":
        return "BANK_ONLY - BRIVA"
    if category == "BFVA":
        return "BANK_ONLY - BFVA"
    return "BANK_ONLY - OTHER"


# ============================================================
# VECTOR / FAST VA EXTRACTION
# ============================================================

def extract_va_series(series, prefix):
    """Ekstraksi VA vectorized. Prefix bank menjadi parameter."""
    regex = rf"({re.escape(prefix)}\d{{5,15}})"
    result = (
        series.astype("string")
        .str.extract(regex, expand=False)
    )
    return result.where(result.notna(), None)


def classify_va_series(series, prefix):
    result = pd.Series(
        "INVALID VA",
        index=series.index,
        dtype="object"
    )

    mask = series.astype("string").str.startswith(prefix, na=False)
    result.loc[mask] = "VALID VA"
    return result


# ============================================================
# BANK CONFIG
# ============================================================

BANK_CONFIGS = {
    "BRIVA": {
        "configured": True,
        "mode": "BRIVA",
        "date_window_days": 0,
        "internal_prefixes": ["57888", "57708"],
        "bank_prefixes": ["57888", "57708"],
        "bank_desc_col": "DESK_TRAN",
        "bank_credit_col": "MUTASI_KREDIT",
        "bank_date_col": "TGL_TRAN",
        "bank_date_format": None,
        "fee_default": 1000,
        "fee_by_prefix": {"57888": 1000, "57708": 1000},
    },
    "BNIVA": {
        "configured": True,
        "mode": "GENERAL",
        "date_window_days": 1,
        "internal_prefixes": ["988765"],
        "bank_prefixes": ["988765"],
        "bank_desc_col": "Description",
        "bank_credit_col": "Credit",
        "bank_date_col": "Post Date",
        "bank_date_format": "%d/%m/%y %H.%M.%S",
        "fee_default": 0,
        "fee_by_prefix": {"988765": 0},
    },
    # Placeholder untuk pengembangan berikutnya
    "BCAVA": {"configured": False},
    "MANDIRIVA": {"configured": False},
    "BSIVA": {"configured": False},
    "MuamalatVA": {"configured": False},
}


# ============================================================
# INTERNAL PREPARATION
# ============================================================

def prepare_internal_dataframe(df_raw, prefixes, fee_by_prefix):
    col_status = find_column(df_raw, ["status", "STATUS"])
    col_desc = find_column(
        df_raw,
        ["keterangan", "KETERANGAN", "description", "DESKRIPSI"]
    )
    col_nominal = find_column(
        df_raw,
        ["nominal", "NOMINAL", "amount", "AMOUNT"]
    )
    col_date = find_column(
        df_raw,
        [
            "tanggal_transfer", "TANGGAL_TRANSFER",
            "tanggal", "TANGGAL",
            "tgl_transfer", "TGL_TRANSFER"
        ]
    )

    df = df_raw.copy()

    status_clean = (
        df[col_status]
        .astype("string")
        .str.strip()
        .str.upper()
    )

    df = df[status_clean == "SUKSES"].copy()

    df["_TANGGAL_DT"] = parse_datetime(df[col_date])
    df = df[df["_TANGGAL_DT"].notna()].copy()

    if df.empty:
        raise ValueError("Tidak ada transaksi FMSS SUKSES dengan tanggal valid.")

    recon_dates = sorted(df["_TANGGAL_DT"].dt.date.dropna().unique())

    # Ekstraksi semua prefix yang relevan.
    df["KODE_VA"] = None
    for prefix in prefixes:
        mask = df["KODE_VA"].isna()
        extracted = extract_va_series(df.loc[mask, col_desc], prefix)
        df.loc[mask, "KODE_VA"] = extracted

    df["JENIS_VA"] = "INVALID VA"
    for prefix in prefixes:
        mask = df["KODE_VA"].astype("string").str.startswith(prefix, na=False)
        df.loc[mask, "JENIS_VA"] = f"VA {prefix}"

    df["NOMINAL_ASLI"] = clean_numeric(df[col_nominal])

    # Fee berdasarkan prefix.
    df["FEE"] = 0
    for prefix, fee in fee_by_prefix.items():
        mask = df["KODE_VA"].astype("string").str.startswith(prefix, na=False)
        df.loc[mask, "FEE"] = fee

    df["EXPECTED_BANK"] = df["NOMINAL_ASLI"] + df["FEE"]

    df_invalid = df[df["KODE_VA"].isna()].copy()
    df_valid = df[df["KODE_VA"].notna()].copy().reset_index(drop=True)

    return df_valid, df_invalid, recon_dates, {
        "col_desc": col_desc,
        "col_nominal": col_nominal,
        "col_date": col_date,
    }


# ============================================================
# BANK PREPARATION
# ============================================================

def extract_bank_va(series, prefixes):
    result = pd.Series(None, index=series.index, dtype="object")
    for prefix in prefixes:
        mask = result.isna()
        extracted = extract_va_series(series.loc[mask], prefix)
        result.loc[mask] = extracted
    return result


def prepare_bank_dataframe(
    df_raw,
    config,
    recon_dates,
    source_bank,
    fee_by_prefix,
):
    col_desc = find_column(df_raw, [config["bank_desc_col"]])
    col_credit = find_column(df_raw, [config["bank_credit_col"]])
    col_date = find_column(df_raw, [config["bank_date_col"]])

    df = df_raw.copy()

    df["_TANGGAL_DT"] = parse_datetime(
        df[col_date],
        config.get("bank_date_format")
    )

    df["_CREDIT_NUM"] = clean_numeric(df[col_credit])

    # Hanya uang masuk.
    df = df[df["_CREDIT_NUM"] > 0].copy()

    # Ekstraksi VA.
    df["KODE_VA"] = extract_bank_va(
        df[col_desc],
        config["bank_prefixes"]
    )

    df["JENIS_VA"] = "INVALID VA"
    for prefix in config["bank_prefixes"]:
        mask = df["KODE_VA"].astype("string").str.startswith(prefix, na=False)
        df.loc[mask, "JENIS_VA"] = f"VA {prefix}"

    # SOURCE / DESCRIPTION.
    df["SOURCE_BANK"] = source_bank
    df["_DESC_VALUE"] = df[col_desc].astype("string").fillna("")
    df["_BANK_TYPE"] = df["_DESC_VALUE"].apply(classify_bank_transaction)

    # ========================================================
    # PERIODE BANK
    #
    # date_window_days = 0  -> persis tanggal FMSS
    # date_window_days = 1  -> H-1 s/d H+1
    #
    # Tanggal BUKAN match key. Ini hanya membatasi data bank yang
    # masuk engine dan kemudian dipakai sebagai audit trail.
    # ========================================================

    target = pd.to_datetime(recon_dates).normalize()
    window = int(config.get("date_window_days", 0))

    if len(target) > 0:
        min_date = target.min() - pd.Timedelta(days=window)
        max_date = target.max() + pd.Timedelta(days=window)
        in_window = (
            df["_TANGGAL_DT"].dt.normalize().between(
                min_date,
                max_date,
                inclusive="both"
            )
        )
    else:
        in_window = pd.Series(False, index=df.index)

    df_window = df[in_window].copy()
    df_outside_window = df[~in_window].copy()

    df_invalid = df_window[df_window["KODE_VA"].isna()].copy()
    df_valid = df_window[df_window["KODE_VA"].notna()].copy().reset_index(drop=True)

    # Data di dalam window tetapi di luar tanggal rekonsiliasi utama.
    target_mask = df_valid["_TANGGAL_DT"].dt.normalize().isin(target)
    df_valid["_IS_TARGET_DATE"] = target_mask

    return df_valid, df_invalid, df_outside_window


# ============================================================
# FAST MATCHING ENGINE
# ============================================================

def fast_match(df_int_valid, df_bank_valid):
    """
    Matching 1-to-1 berbasis dictionary + deque.

    LOGIC UTAMA TETAP:
      KODE_VA sama
      EXPECTED_BANK == MUTASI_KREDIT

    Tanggal TIDAK menjadi match key.
    Ini penting untuk cutoff / posting H-1 atau H+1.
    """

    bank_records = df_bank_valid.to_dict("records")

    bank_index = defaultdict(deque)

    for idx, bank_row in enumerate(bank_records):
        key = (
            str(bank_row["KODE_VA"]),
            float(bank_row["_CREDIT_NUM"])
        )
        bank_index[key].append(idx)

    matched_bank_indexes = set()
    matched = []
    unmatched_internal = []

    for int_row in df_int_valid.to_dict("records"):
        key = (
            str(int_row["KODE_VA"]),
            float(int_row["EXPECTED_BANK"])
        )

        queue = bank_index.get(key)

        if queue:
            bank_idx = queue.popleft()
            bank_row = bank_records[bank_idx]
            matched_bank_indexes.add(bank_idx)

            record = int_row.copy()
            record["MATCH_MUTASI_KREDIT"] = bank_row["_CREDIT_NUM"]
            record["MATCH_DESK_TRAN"] = bank_row.get("_DESC_VALUE", "")
            record["SOURCE_BANK"] = bank_row.get("SOURCE_BANK", "")
            record["BANK_TYPE"] = bank_row.get("_BANK_TYPE", "")
            record["BANK_TANGGAL"] = bank_row.get("_TANGGAL_DT")

            int_date = record.get("_TANGGAL_DT")
            bank_date = bank_row.get("_TANGGAL_DT")

            if pd.notna(int_date) and pd.notna(bank_date):
                record["SELISIH_HARI"] = (
                    pd.Timestamp(bank_date).normalize()
                    - pd.Timestamp(int_date).normalize()
                ).days
            else:
                record["SELISIH_HARI"] = None

            if record["SELISIH_HARI"] == 0:
                record["MATCH_DATE_STATUS"] = "SAME DAY"
            elif record["SELISIH_HARI"] == -1:
                record["MATCH_DATE_STATUS"] = "BANK H-1"
            elif record["SELISIH_HARI"] == 1:
                record["MATCH_DATE_STATUS"] = "BANK H+1"
            else:
                record["MATCH_DATE_STATUS"] = "BANK DI LUAR ±1 HARI"

            record["STATUS_MATCH"] = "MATCHED"
            matched.append(record)

        else:
            record = int_row.copy()
            record["STATUS_MATCH"] = "FMSS_ONLY"
            unmatched_internal.append(record)

    # Sisa bank = bank-only.
    unmatched_bank = []
    for idx, bank_row in enumerate(bank_records):
        if idx in matched_bank_indexes:
            continue

        record = bank_row.copy()
        record["STATUS_MATCH"] = classify_issue_bank(
            bank_row.get("_DESC_VALUE", "")
        )
        unmatched_bank.append(record)

    return (
        pd.DataFrame(matched),
        pd.DataFrame(unmatched_internal),
        pd.DataFrame(unmatched_bank),
    )


# ============================================================
# SUMMARY HELPER
# ============================================================

def build_summary(df_matched, df_fmss_only, df_bank_only, df_outside, df_invalid_int, df_invalid_bnk):
    return {
        "matched_count": len(df_matched),
        "fmss_only_count": len(df_fmss_only),
        "bank_only_count": len(df_bank_only),
        "out_of_period_count": len(df_outside),
        "invalid_int_count": len(df_invalid_int),
        "invalid_bnk_count": len(df_invalid_bnk),
        "matched_nominal": df_matched["NOMINAL_ASLI"].sum() if not df_matched.empty else 0,
        "fmss_only_nominal": df_fmss_only["NOMINAL_ASLI"].sum() if not df_fmss_only.empty else 0,
        "bank_only_nominal": df_bank_only["_CREDIT_NUM"].sum() if not df_bank_only.empty else 0,
        "out_of_period_nominal": df_outside["_CREDIT_NUM"].sum() if not df_outside.empty else 0,
    }


# ============================================================
# UI - PILIH BANK
# ============================================================

st.subheader("1. Pengaturan Data")

opsi_bank = ["", "BRIVA", "BNIVA", "BCAVA", "MANDIRIVA", "BSIVA", "MuamalatVA"]
pilihan_bank = st.selectbox("Pilih Bank Sumber Mutasi:", opsi_bank)

if pilihan_bank != st.session_state.pilihan_bank_terakhir:
    for key in [
        "sudah_diproses", "df_matched", "df_selisih_int",
        "df_selisih_bnk", "df_out_of_period", "df_invalid_int",
        "df_invalid_bnk", "recon_dates", "summary"
    ]:
        if key in ["recon_dates"]:
            st.session_state[key] = []
        elif key == "summary":
            st.session_state[key] = {}
        elif key == "sudah_diproses":
            st.session_state[key] = False
        else:
            st.session_state[key] = pd.DataFrame()

    st.session_state.pilihan_bank_terakhir = pilihan_bank


# ============================================================
# UI - UPLOAD
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
        st.markdown(f"### 🏦 Mutasi {pilihan_bank or 'Bank'}")
        file_bnk_general = st.file_uploader(
            f"Upload mutasi {pilihan_bank or 'Bank'}",
            type=["csv", "xlsx"],
            key="bank_general"
        )


# ============================================================
# FEE CONFIG
# ============================================================

fee_57888 = 1000
fee_57708 = 1000

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

    st.caption("Rumus: Nominal FMSS + Fee = Nominal mutasi bank.")

elif pilihan_bank == "BNIVA":
    st.subheader("3. Konfigurasi Fee")
    st.info("BNIVA terdeteksi tanpa fee pada sample: Nominal FMSS = Nominal bank.")


# ============================================================
# PROCESS BUTTON
# ============================================================

can_process = False

if pilihan_bank == "BRIVA":
    can_process = bool(file_int and file_bnk_57888 and file_bnk_57708)
elif pilihan_bank == "BNIVA":
    can_process = bool(file_int and file_bnk_general)

if can_process:
    st.divider()

    if st.button(
        f"🚀 Mulai Croscek Data {pilihan_bank}",
        type="primary",
        use_container_width=True
    ):
        st.session_state.sudah_diproses = False

        try:
            with st.spinner(f"Sedang memproses rekonsiliasi {pilihan_bank}..."):

                # ------------------------------------------------
                # FMSS
                # ------------------------------------------------
                df_int_raw = read_uploaded_file(file_int)

                if pilihan_bank == "BRIVA":
                    fee_by_prefix = {
                        "57888": fee_57888,
                        "57708": fee_57708,
                    }
                else:
                    fee_by_prefix = {
                        "988765": 0,
                    }

                config = BANK_CONFIGS[pilihan_bank]

                df_int_valid, df_invalid_int, recon_dates, internal_meta = prepare_internal_dataframe(
                    df_int_raw,
                    config["internal_prefixes"],
                    fee_by_prefix,
                )

                st.session_state.recon_dates = recon_dates

                # ------------------------------------------------
                # BANK
                # ------------------------------------------------
                bank_frames = []
                outside_frames = []

                if pilihan_bank == "BRIVA":
                    df_57888_raw = read_uploaded_file(file_bnk_57888)
                    df_57708_raw = read_uploaded_file(file_bnk_57708)

                    config_57888 = dict(config)
                    config_57888["bank_prefixes"] = ["57888"]
                    config_57888["bank_date_col"] = "TGL_TRAN"

                    config_57708 = dict(config)
                    config_57708["bank_prefixes"] = ["57708"]
                    config_57708["bank_date_col"] = "TGL_TRAN"

                    b1, i1, o1 = prepare_bank_dataframe(
                        df_57888_raw,
                        config_57888,
                        recon_dates,
                        "BRIVA FASTPAY 57888",
                        {"57888": fee_57888},
                    )

                    b2, i2, o2 = prepare_bank_dataframe(
                        df_57708_raw,
                        config_57708,
                        recon_dates,
                        "BRIVA RAJABILLER 57708",
                        {"57708": fee_57708},
                    )

                    bank_frames.extend([b1, b2])
                    invalid_bank_frames = [i1, i2]
                    outside_frames.extend([o1, o2])

                elif pilihan_bank == "BNIVA":
                    df_bnk_raw = read_uploaded_file(file_bnk_general)

                    b1, i1, o1 = prepare_bank_dataframe(
                        df_bnk_raw,
                        config,
                        recon_dates,
                        "BNIVA",
                        {"988765": 0},
                    )

                    bank_frames.append(b1)
                    invalid_bank_frames = [i1]
                    outside_frames.append(o1)

                else:
                    raise ValueError(
                        f"Bank {pilihan_bank} belum dikonfigurasi."
                    )

                df_bank_valid = pd.concat(
                    bank_frames,
                    ignore_index=True
                ) if bank_frames else pd.DataFrame()

                df_invalid_bnk = pd.concat(
                    invalid_bank_frames,
                    ignore_index=True
                ) if invalid_bank_frames else pd.DataFrame()

                df_outside = pd.concat(
                    outside_frames,
                    ignore_index=True
                ) if outside_frames else pd.DataFrame()

                # ------------------------------------------------
                # MATCHING
                # ------------------------------------------------
                (
                    df_matched,
                    df_selisih_int,
                    df_selisih_bnk
                ) = fast_match(
                    df_int_valid,
                    df_bank_valid
                )

                # ------------------------------------------------
                # IMPORTANT:
                # Untuk BNIVA, transaksi bank yang masih tersisa di
                # H-1/H+1 tetapi bukan tanggal rekonsiliasi utama
                # tidak dianggap Issue Bank. Mereka dipindah ke
                # OUT_OF_PERIOD agar tidak menggelembungkan issue.
                # ------------------------------------------------
                if not df_selisih_bnk.empty:
                    target_dates = pd.to_datetime(recon_dates).normalize()
                    bank_dates = pd.to_datetime(
                        df_selisih_bnk["_TANGGAL_DT"],
                        errors="coerce"
                    ).dt.normalize()

                    is_target = bank_dates.isin(target_dates)

                    df_bank_target = df_selisih_bnk[is_target].copy()
                    df_bank_near = df_selisih_bnk[~is_target].copy()

                    if not df_bank_near.empty:
                        df_outside = pd.concat(
                            [df_outside, df_bank_near],
                            ignore_index=True
                        )

                    df_selisih_bnk = df_bank_target.reset_index(drop=True)

                # ------------------------------------------------
                # SUMMARY
                # ------------------------------------------------
                summary = build_summary(
                    df_matched,
                    df_selisih_int,
                    df_selisih_bnk,
                    df_outside,
                    df_invalid_int,
                    df_invalid_bnk,
                )

                st.session_state.df_matched = df_matched
                st.session_state.df_selisih_int = df_selisih_int
                st.session_state.df_selisih_bnk = df_selisih_bnk
                st.session_state.df_out_of_period = df_outside
                st.session_state.df_invalid_int = df_invalid_int
                st.session_state.df_invalid_bnk = df_invalid_bnk
                st.session_state.summary = summary
                st.session_state.sudah_diproses = True

        except Exception as e:
            st.session_state.sudah_diproses = False
            st.error("❌ Terjadi kesalahan saat memproses data.")
            st.exception(e)


# ============================================================
# RESULTS
# ============================================================

if st.session_state.sudah_diproses:
    df_matched = st.session_state.df_matched
    df_selisih_int = st.session_state.df_selisih_int
    df_selisih_bnk = st.session_state.df_selisih_bnk
    df_outside = st.session_state.df_out_of_period
    df_invalid_int = st.session_state.df_invalid_int
    df_invalid_bnk = st.session_state.df_invalid_bnk
    summary = st.session_state.summary

    st.divider()
    st.subheader(f"🎯 Ringkasan Rekonsiliasi {pilihan_bank}")
    st.caption(
        f"Periode FMSS: **{safe_date_string(st.session_state.recon_dates)}**"
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("✅ Matched Sempurna", f"{summary['matched_count']:,} Trx")
    m2.metric("⚠️ Issue FMSS", f"{summary['fmss_only_count']:,} Trx")
    m3.metric("⚠️ Issue Bank", f"{summary['bank_only_count']:,} Trx")
    m4.metric("📅 Di Luar Periode", f"{summary['out_of_period_count']:,} Trx")

    total_fmss_valid = len(df_matched) + len(df_selisih_int)
    total_bank_valid = len(df_matched) + len(df_selisih_bnk)

    fmss_match_rate = (
        len(df_matched) / total_fmss_valid * 100
        if total_fmss_valid else 0
    )
    bank_match_rate = (
        len(df_matched) / total_bank_valid * 100
        if total_bank_valid else 0
    )

    r1, r2 = st.columns(2)
    r1.metric("📈 Match Rate FMSS", f"{fmss_match_rate:.4f}%")
    r2.metric("📈 Match Rate Bank", f"{bank_match_rate:.4f}%")

    st.subheader("💰 Ringkasan Nominal")
    n1, n2, n3, n4 = st.columns(4)
    n1.metric("Matched", format_rupiah(summary["matched_nominal"]))
    n2.metric("Issue FMSS", format_rupiah(summary["fmss_only_nominal"]))
    n3.metric("Issue Bank", format_rupiah(summary["bank_only_nominal"]))
    n4.metric("Di Luar Periode", format_rupiah(summary["out_of_period_nominal"]))

    # --------------------------------------------------------
    # MATCH DATE SHIFT
    # --------------------------------------------------------
    if not df_matched.empty and "MATCH_DATE_STATUS" in df_matched.columns:
        st.divider()
        st.subheader("🕒 Audit Cutoff / Perbedaan Tanggal")
        shift_counts = (
            df_matched["MATCH_DATE_STATUS"]
            .value_counts()
            .rename_axis("STATUS")
            .reset_index(name="JUMLAH")
        )
        st.dataframe(
            shift_counts,
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # ISSUES
    # --------------------------------------------------------
    st.divider()
    col_issue1, col_issue2 = st.columns(2)

    with col_issue1:
        st.subheader("🚨 Issue FMSS")
        if not df_selisih_int.empty:
            cols = [
                "KODE_VA", "JENIS_VA", "NOMINAL_ASLI",
                "FEE", "EXPECTED_BANK", "STATUS_MATCH"
            ]
            cols = [c for c in cols if c in df_selisih_int.columns]
            st.dataframe(
                df_selisih_int[cols],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.success("Tidak ada issue FMSS.")

    with col_issue2:
        st.subheader("🚨 Issue Bank")
        if not df_selisih_bnk.empty:
            cols = [
                "KODE_VA", "JENIS_VA", "_CREDIT_NUM",
                "_TANGGAL_DT", "SOURCE_BANK", "STATUS_MATCH"
            ]
            cols = [c for c in cols if c in df_selisih_bnk.columns]
            st.dataframe(
                df_selisih_bnk[cols],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.success("Tidak ada issue Bank pada tanggal rekonsiliasi.")

    # --------------------------------------------------------
    # OUT OF PERIOD
    # --------------------------------------------------------
    if not df_outside.empty:
        st.divider()
        with st.expander(
            f"📅 Transaksi Bank Di Luar Periode Utama ({len(df_outside):,} Trx)",
            expanded=False
        ):
            cols = [
                "KODE_VA", "JENIS_VA", "_CREDIT_NUM",
                "_TANGGAL_DT", "SOURCE_BANK", "STATUS_MATCH"
            ]
            cols = [c for c in cols if c in df_outside.columns]
            st.dataframe(
                df_outside[cols],
                use_container_width=True,
                hide_index=True
            )

    # --------------------------------------------------------
    # INVALID VA
    # --------------------------------------------------------
    st.divider()
    with st.expander("⚠️ Transaksi dengan VA Tidak Teridentifikasi"):
        iv1, iv2 = st.columns(2)

        with iv1:
            st.markdown("### FMSS Invalid VA")
            if not df_invalid_int.empty:
                st.dataframe(
                    df_invalid_int,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.success("Tidak ada FMSS invalid VA.")

        with iv2:
            st.markdown("### Bank Invalid VA")
            if not df_invalid_bnk.empty:
                st.dataframe(
                    df_invalid_bnk,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.success("Tidak ada Bank invalid VA.")

    # --------------------------------------------------------
    # DOWNLOAD
    # --------------------------------------------------------
    st.divider()
    st.subheader("📥 Download Laporan")

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_export = pd.DataFrame({
            "METRIC": [
                "Bank",
                "Periode FMSS",
                "Matched",
                "FMSS Only",
                "Bank Only",
                "Out of Period Bank",
                "FMSS Invalid VA",
                "Bank Invalid VA",
                "Match Rate FMSS",
                "Match Rate Bank",
                "Nominal Matched",
                "Nominal FMSS Only",
                "Nominal Bank Only",
                "Nominal Out of Period Bank",
            ],
            "VALUE": [
                pilihan_bank,
                safe_date_string(st.session_state.recon_dates),
                summary["matched_count"],
                summary["fmss_only_count"],
                summary["bank_only_count"],
                summary["out_of_period_count"],
                summary["invalid_int_count"],
                summary["invalid_bnk_count"],
                f"{fmss_match_rate:.4f}%",
                f"{bank_match_rate:.4f}%",
                summary["matched_nominal"],
                summary["fmss_only_nominal"],
                summary["bank_only_nominal"],
                summary["out_of_period_nominal"],
            ]
        })
        summary_export.to_excel(writer, sheet_name="SUMMARY", index=False)

        exports = [
            (df_matched, "MATCHED_OK", "Tidak ada data matched."),
            (df_selisih_int, "ISSUE_FMSS", "Tidak ada issue FMSS."),
            (df_selisih_bnk, "ISSUE_BANK", "Tidak ada issue Bank."),
            (df_outside, "OUT_OF_PERIOD_BANK", "Tidak ada transaksi di luar periode."),
            (df_invalid_int, "INVALID_FMSS", "Tidak ada FMSS invalid VA."),
            (df_invalid_bnk, "INVALID_BANK", "Tidak ada Bank invalid VA."),
        ]

        for df_export, sheet, empty_msg in exports:
            if not df_export.empty:
                clean_export = df_export.drop(
                    columns=["_TANGGAL_ONLY", "_IS_TARGET_DATE"],
                    errors="ignore"
                ).copy()
                clean_export.to_excel(writer, sheet_name=sheet, index=False)
            else:
                pd.DataFrame({"INFO": [empty_msg]}).to_excel(
                    writer,
                    sheet_name=sheet,
                    index=False
                )

    output.seek(0)

    st.download_button(
        label="📥 Download Laporan Lengkap (.xlsx)",
        data=output.getvalue(),
        file_name=(
            f"Laporan_Rekonsiliasi_{pilihan_bank}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

elif pilihan_bank == "BRIVA":
    st.info("💡 Upload 3 file: FMSS, Mutasi BRIVA 57888, dan Mutasi BRIVA 57708.")
elif pilihan_bank == "BNIVA":
    st.info("💡 Upload 2 file: FMSS dan Mutasi BNIVA.")
elif pilihan_bank:
    st.warning(f"🚧 Modul {pilihan_bank} belum dikonfigurasi.")
