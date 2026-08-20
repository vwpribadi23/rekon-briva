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

    # Exact match
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

        # Tetap menggunakan logic pembacaan CSV
        # yang sama dengan script sebelumnya.
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

        return pd.to_numeric(
            series,
            errors="coerce"
        ).fillna(0)

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

    match = re.search(
        r"(57(?:888|708)\d{5,15})",
        text
    )

    if match:
        return match.group(1)

    return None


def classify_va(va):

    if (
        pd.isna(va)
        or va is None
        or str(va).strip() == ""
    ):
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

        return sorted_dates[0].strftime(
            "%d %B %Y"
        )

    return (
        f"{sorted_dates[0].strftime('%d %B %Y')} "
        f"s/d {sorted_dates[-1].strftime('%d %B %Y')}"
    )


def format_rupiah(value):

    try:
        value = float(value)
    except:
        value = 0

    return "Rp {:,.0f}".format(
        value
    ).replace(",", ".")


def classify_issue_bank(description):

    category = classify_bank_transaction(
        description
    )

    if category in [
        "ATM / MANUAL",
        "TRANSFER / MANUAL"
    ]:
        return "BANK_ONLY - MANUAL/ATM"

    if category == "BRIVA":
        return "BANK_ONLY - BRIVA"

    if category == "BFVA":
        return "BANK_ONLY - BFVA"

    return "BANK_ONLY - OTHER"


# ============================================================
# FAST VA EXTRACTION
# ============================================================

VA_REGEX = r"(57(?:888|708)\d{5,15})"


def extract_va_series(series):
    """
    Versi vectorized dari extract_va().
    Hasil dibuat konsisten dengan logic sebelumnya.
    """

    result = (
        series.astype("string")
        .str.extract(
            VA_REGEX,
            expand=False
        )
    )

    return result.where(
        result.notna(),
        None
    )


def classify_va_series(series):

    result = pd.Series(
        "INVALID VA",
        index=series.index,
        dtype="object"
    )

    mask_57888 = (
        series.astype("string")
        .str.startswith("57888", na=False)
    )

    mask_57708 = (
        series.astype("string")
        .str.startswith("57708", na=False)
    )

    result.loc[mask_57888] = "BRIVA FASTPAY"
    result.loc[mask_57708] = "BRIVA RAJABILLER"

    return result


# ============================================================
# BRIVA ENGINE - COMPREHENSIVE DATE / CUTOFF HANDLING
# ============================================================
# Catatan:
# - fast_match() BRIVA legacy di bawah TIDAK diubah.
# - Jalur BRIVA baru ini hanya memperbaiki temuan rekonsiliasi:
#   1) timestamp FMSS campuran (dengan / tanpa microsecond),
#   2) mutasi D+1 boleh menjadi search pool untuk cutoff,
#   3) unmatched D+1 tidak dihitung sebagai Issue Bank D,
#   4) matching tetap VA + EXPECTED_BANK, 1-to-1.
#   5) credit leg reversal/net-zero BRI dikeluarkan sebelum matching.
# - BNIVA, MANDIRIVA, dan engine tampilan dashboard tidak berubah.


def parse_briva_datetime(series):
    """
    Parser khusus BRIVA yang tahan timestamp campuran.

    Contoh yang harus sama-sama valid:
        2026-08-19 23:58:49.339228
        2026-08-19 23:00:04

    pd.to_datetime(series) pada beberapa versi pandas dapat mengunci
    satu format dari baris awal sehingga baris tanpa microsecond menjadi NaT.
    Karena itu parsing gagal akan dicoba ulang per nilai.
    """

    # Pandas modern: format='mixed' adalah pilihan paling tepat.
    try:
        parsed = pd.to_datetime(
            series,
            errors="coerce",
            format="mixed"
        )
    except (TypeError, ValueError):
        parsed = pd.to_datetime(
            series,
            errors="coerce"
        )

    # Fallback untuk versi pandas yang belum mendukung format='mixed'
    # atau apabila masih ada nilai valid yang gagal diparse secara vectorized.
    source = pd.Series(
        series,
        index=getattr(series, "index", None)
    )

    missing_mask = (
        parsed.isna()
        & source.notna()
        & source.astype(str).str.strip().ne("")
    )

    if missing_mask.any():
        reparsed = source.loc[missing_mask].apply(
            lambda value: pd.to_datetime(
                value,
                errors="coerce"
            )
        )

        parsed.loc[missing_mask] = reparsed

    return parsed


def build_briva_search_dates(recon_dates):
    """
    Search pool BRIVA untuk FMSS tanggal D:
        D   = transaksi normal / retry yang sukses pada hari D
        D+1 = antisipasi cutoff bank

    D+1 hanya search pool. Sisa mutasi D+1 tidak boleh menjadi Issue Bank D.
    """

    result = set()

    for value in recon_dates:
        base_date = pd.Timestamp(value).date()
        result.add(base_date)
        result.add(
            base_date + timedelta(days=1)
        )

    return sorted(result)


BRIVA_REVERSAL_WINDOW_SECONDS = 120


def _briva_amount_key(value):
    """
    Key nominal BRIVA untuk kebutuhan deteksi reversal.
    Pembulatan ke rupiah penuh konsisten dengan data mutasi BRI.
    """

    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return 0


def detect_briva_reversal_credit_mask(df):
    """
    Mendeteksi credit leg BRIVA yang sudah direversal / dibalik bank.

    Prinsip konservatif:
    1) VA harus sama.
    2) Description harus sama persis.
    3) Jika SEQ tersedia, SEQ harus sama.
    4) Total credit dan debit dalam satu event harus sama (net zero).
    5) Seluruh event terjadi dalam <= 120 detik.

    Pada file BRI 20 Agustus yang diuji, pola reversal contohnya:
        Credit nominal utama
        Debit fee / adjustment
        Credit fee / adjustment
        Debit nominal utama

    Total credit == total debit, sehingga uang bersih yang bertahan di
    rekening adalah nol. Semua baris CREDIT dari event tersebut tidak boleh
    dipakai untuk matching FMSS maupun dihitung sebagai Issue Bank.

    Jika kolom debit tidak tersedia, fungsi tidak mengubah perilaku lama.
    """

    result = pd.Series(
        False,
        index=df.index,
        dtype=bool
    )

    required_cols = {
        "_TANGGAL_DT",
        "_CREDIT_NUM",
        "_DEBIT_NUM",
        "KODE_VA",
        "_DESC_VALUE"
    }

    if not required_cols.issubset(df.columns):
        return result

    work = df[
        df["KODE_VA"].notna()
        & (
            (df["_CREDIT_NUM"] > 0)
            | (df["_DEBIT_NUM"] > 0)
        )
        & df["_TANGGAL_DT"].notna()
    ].copy()

    if work.empty:
        return result

    # Normalisasi key agar grouping tidak terganggu tipe data campuran.
    work["_REVERSAL_VA_KEY"] = (
        work["KODE_VA"]
        .astype(str)
        .str.strip()
    )

    work["_REVERSAL_DESC_KEY"] = (
        work["_DESC_VALUE"]
        .astype(str)
        .str.strip()
    )

    work["_REVERSAL_DATE_KEY"] = (
        work["_TANGGAL_DT"]
        .dt.date
    )

    # SEQ bersifat optional. Jika file BRI menyediakan SEQ, kita pakai
    # sebagai penguat event ID agar deteksi reversal makin konservatif.
    has_seq = (
        "_BRIVA_SEQ_KEY" in work.columns
        and work["_BRIVA_SEQ_KEY"].notna().any()
    )

    if has_seq:
        group_cols = [
            "_REVERSAL_VA_KEY",
            "_REVERSAL_DESC_KEY",
            "_BRIVA_SEQ_KEY",
            "_REVERSAL_DATE_KEY"
        ]

        for _, group in work.groupby(
            group_cols,
            dropna=False,
            sort=False
        ):
            total_credit = float(
                group["_CREDIT_NUM"].sum()
            )

            total_debit = float(
                group["_DEBIT_NUM"].sum()
            )

            if total_credit <= 0 or total_debit <= 0:
                continue

            # Net zero sampai toleransi < Rp1.
            if abs(total_credit - total_debit) >= 0.5:
                continue

            time_span = (
                group["_TANGGAL_DT"].max()
                - group["_TANGGAL_DT"].min()
            ).total_seconds()

            if time_span > BRIVA_REVERSAL_WINDOW_SECONDS:
                continue

            credit_indexes = group.index[
                group["_CREDIT_NUM"] > 0
            ]

            result.loc[credit_indexes] = True

        return result

    # --------------------------------------------------------
    # FALLBACK JIKA SEQ TIDAK TERSEDIA
    # --------------------------------------------------------
    # Pair credit-debit one-to-one berdasarkan VA + description + nominal,
    # dalam window waktu yang sama. Ini sengaja lebih ketat daripada hanya
    # VA + nominal agar transaksi reguler tidak salah dianggap reversal.

    debit_pool = defaultdict(list)

    for idx, row in work[
        work["_DEBIT_NUM"] > 0
    ].iterrows():
        key = (
            row["_REVERSAL_VA_KEY"],
            row["_REVERSAL_DESC_KEY"],
            row["_REVERSAL_DATE_KEY"],
            _briva_amount_key(
                row["_DEBIT_NUM"]
            )
        )

        debit_pool[key].append(
            (idx, row["_TANGGAL_DT"])
        )

    used_debits = set()

    for credit_idx, row in work[
        work["_CREDIT_NUM"] > 0
    ].sort_values(
        "_TANGGAL_DT",
        kind="stable"
    ).iterrows():
        key = (
            row["_REVERSAL_VA_KEY"],
            row["_REVERSAL_DESC_KEY"],
            row["_REVERSAL_DATE_KEY"],
            _briva_amount_key(
                row["_CREDIT_NUM"]
            )
        )

        candidates = [
            item
            for item in debit_pool.get(key, [])
            if item[0] not in used_debits
        ]

        if not candidates:
            continue

        nearest_idx, nearest_dt = min(
            candidates,
            key=lambda item: abs(
                (
                    item[1]
                    - row["_TANGGAL_DT"]
                ).total_seconds()
            )
        )

        delta_seconds = abs(
            (
                nearest_dt
                - row["_TANGGAL_DT"]
            ).total_seconds()
        )

        if delta_seconds <= BRIVA_REVERSAL_WINDOW_SECONDS:
            result.loc[credit_idx] = True
            used_debits.add(
                nearest_idx
            )

    return result


def prepare_briva_bank_dataframe(
    uploaded_file,
    recon_dates,
    source_bank
):
    """
    Load dan normalisasi file bank khusus BRIVA.

    Logic field/VA/nominal tetap mengikuti prepare_bank_dataframe() legacy.
    Enhancement terbatas:
        - parser tanggal BRIVA lebih robust,
        - bank D+1 ikut dibaca sebagai search pool cutoff,
        - credit leg reversal / net-zero dikeluarkan sebelum matching.

    Reversal hanya dideteksi jika file menyediakan kolom MUTASI_DEBET/DEBET.
    Jika kolom debit tidak tersedia, perilaku tetap sama seperti versi lama.
    """

    df = read_uploaded_file(
        uploaded_file
    )

    col_credit = find_column(
        df,
        [
            "MUTASI_KREDIT",
            "mutasi_kredit",
            "KREDIT",
            "kredit"
        ]
    )

    # Debit bersifat optional agar kompatibel dengan file BRIVA ringkas
    # yang sebelumnya hanya diwajibkan TGL_TRAN, DESK_TRAN, MUTASI_KREDIT.
    col_debit = find_column(
        df,
        [
            "MUTASI_DEBET",
            "mutasi_debet",
            "DEBET",
            "debet",
            "DEBIT",
            "debit"
        ],
        required=False
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
            "tanggal"
        ]
    )

    col_seq = find_column(
        df,
        [
            "SEQ",
            "seq",
            "SEQUENCE",
            "sequence"
        ],
        required=False
    )

    df = df.copy()

    # --------------------------------------------------------
    # DATE - parser khusus BRIVA
    # --------------------------------------------------------

    df["_TANGGAL_DT"] = parse_briva_datetime(
        df[col_date]
    )

    # --------------------------------------------------------
    # CREDIT / DEBIT
    # --------------------------------------------------------

    df["_CREDIT_NUM"] = clean_numeric(
        df[col_credit]
    )

    if col_debit is not None:
        df["_DEBIT_NUM"] = clean_numeric(
            df[col_debit]
        )
    else:
        df["_DEBIT_NUM"] = 0.0

    # --------------------------------------------------------
    # FILTER TANGGAL D + D+1
    # --------------------------------------------------------

    search_dates = build_briva_search_dates(
        recon_dates
    )

    search_datetime = pd.to_datetime(
        search_dates
    )

    df["_TANGGAL_ONLY"] = (
        df["_TANGGAL_DT"]
        .dt.normalize()
    )

    df = df[
        df["_TANGGAL_ONLY"].isin(
            search_datetime
        )
    ].copy()

    # --------------------------------------------------------
    # BANK TYPE - logic legacy
    # --------------------------------------------------------

    df["_BANK_TYPE"] = (
        df[col_desc]
        .apply(classify_bank_transaction)
    )

    # --------------------------------------------------------
    # VA - logic legacy
    # --------------------------------------------------------

    df["KODE_VA"] = extract_va_series(
        df[col_desc]
    )

    df["JENIS_VA"] = classify_va_series(
        df["KODE_VA"]
    )

    # --------------------------------------------------------
    # SOURCE / AUDIT FIELDS
    # --------------------------------------------------------

    df["SOURCE_BANK"] = source_bank

    df["_DESC_VALUE"] = (
        df[col_desc]
        .astype(str)
    )

    if col_seq is not None:
        df["_BRIVA_SEQ_KEY"] = (
            df[col_seq]
            .astype("string")
            .str.strip()
        )
    else:
        df["_BRIVA_SEQ_KEY"] = pd.NA

    # --------------------------------------------------------
    # REVERSAL / NET-ZERO DETECTION
    # --------------------------------------------------------
    # Dilakukan sebelum filter CREDIT > 0 agar sisi debit dari event reversal
    # masih tersedia untuk membuktikan bahwa credit tersebut sudah dibalik.

    df["_IS_REVERSAL_CREDIT"] = (
        detect_briva_reversal_credit_mask(
            df
        )
    )

    df["_REVERSAL_STATUS"] = ""
    df.loc[
        df["_IS_REVERSAL_CREDIT"],
        "_REVERSAL_STATUS"
    ] = "REVERSAL_NET_ZERO_EXCLUDED"

    # --------------------------------------------------------
    # HANYA UANG MASUK YANG MASIH VALID
    # --------------------------------------------------------

    df = df[
        (df["_CREDIT_NUM"] > 0)
        & (~df["_IS_REVERSAL_CREDIT"])
    ].copy()

    return df


def fast_match_briva(
    df_int_valid,
    df_bank_valid,
    recon_dates
):
    """
    Matching BRIVA komprehensif dengan rule bisnis lama:
        KODE_VA exact
        EXPECTED_BANK exact
        1-to-1

    Enhancement terbatas:
        - Bank tanggal D diprioritaskan sebelum D+1.
        - D+1 dapat menyelesaikan FMSS cutoff.
        - Sisa Bank D+1 tidak dihitung sebagai Issue Bank tanggal D.

    Tidak ada perubahan pada rumus fee maupun klasifikasi BRIVA.
    """

    target_dates = {
        pd.Timestamp(value).date()
        for value in recon_dates
    }

    # Copy agar dataframe asli tidak berubah.
    bank_work = df_bank_valid.copy()

    if not bank_work.empty:
        bank_work["_BRIVA_ORIGINAL_ORDER"] = range(
            len(bank_work)
        )

        bank_date = (
            bank_work["_TANGGAL_DT"]
            .dt.date
        )

        # Tanggal target D selalu dicoba lebih dulu daripada D+1.
        bank_work["_BRIVA_DATE_PRIORITY"] = (
            ~bank_date.isin(target_dates)
        ).astype(int)

        bank_work = bank_work.sort_values(
            by=[
                "_BRIVA_DATE_PRIORITY",
                "_TANGGAL_DT",
                "_BRIVA_ORIGINAL_ORDER"
            ],
            kind="stable",
            na_position="last"
        ).reset_index(drop=True)

    bank_records = (
        bank_work
        .to_dict("records")
    )

    bank_index = defaultdict(
        deque
    )

    for idx, bank_row in enumerate(
        bank_records
    ):
        key = (
            str(bank_row["KODE_VA"]),
            float(bank_row["_CREDIT_NUM"])
        )

        bank_index[key].append(
            idx
        )

    matched_bank_indexes = set()
    matched = []
    unmatched_internal = []

    int_records = (
        df_int_valid
        .to_dict("records")
    )

    for int_row in int_records:
        key = (
            str(int_row["KODE_VA"]),
            float(int_row["EXPECTED_BANK"])
        )

        queue = bank_index.get(
            key
        )

        if queue:
            bank_idx = queue.popleft()
            bank_row = bank_records[bank_idx]

            matched_bank_indexes.add(
                bank_idx
            )

            record = int_row.copy()

            record["MATCH_MUTASI_KREDIT"] = (
                bank_row["_CREDIT_NUM"]
            )

            record["MATCH_DESK_TRAN"] = (
                bank_row.get(
                    "_DESC_VALUE",
                    ""
                )
            )

            record["SOURCE_BANK"] = (
                bank_row.get(
                    "SOURCE_BANK",
                    ""
                )
            )

            record["BANK_TYPE"] = (
                bank_row.get(
                    "_BANK_TYPE",
                    ""
                )
            )

            record["STATUS_MATCH"] = (
                "MATCHED"
            )

            # Metadata audit tambahan; tidak mengubah dashboard.
            bank_dt = bank_row.get(
                "_TANGGAL_DT"
            )

            fmss_dt = int_row.get(
                "_TANGGAL_DT"
            )

            record["MATCH_BANK_DATETIME"] = bank_dt

            if (
                pd.notna(bank_dt)
                and pd.notna(fmss_dt)
            ):
                bank_date_value = pd.Timestamp(
                    bank_dt
                ).date()
                fmss_date_value = pd.Timestamp(
                    fmss_dt
                ).date()

                if bank_date_value == fmss_date_value:
                    record["DATE_RELATION"] = "SAME_DAY"
                elif bank_date_value == (
                    fmss_date_value + timedelta(days=1)
                ):
                    record["DATE_RELATION"] = "H+1_CUTOFF"
                else:
                    record["DATE_RELATION"] = "OTHER"

            matched.append(
                record
            )

        else:
            record = int_row.copy()
            record["STATUS_MATCH"] = (
                "FMSS_ONLY"
            )

            unmatched_internal.append(
                record
            )

    # --------------------------------------------------------
    # ISSUE BANK
    # --------------------------------------------------------
    # Hanya sisa mutasi bank tanggal TARGET D yang masuk Issue Bank.
    # D+1 adalah search pool cutoff dan tidak boleh memperbesar issue D.

    unmatched_bank = []

    for idx, bank_row in enumerate(
        bank_records
    ):
        if idx in matched_bank_indexes:
            continue

        bank_dt = bank_row.get(
            "_TANGGAL_DT"
        )

        if pd.isna(bank_dt):
            continue

        bank_date_value = pd.Timestamp(
            bank_dt
        ).date()

        if bank_date_value not in target_dates:
            continue

        record = bank_row.copy()

        record["STATUS_MATCH"] = (
            classify_issue_bank(
                bank_row.get(
                    "_BANK_TYPE",
                    ""
                )
            )
        )

        record["MATCH_METHOD"] = (
            "NO_FMSS_MATCH_TARGET_DATE"
        )

        record["MATCH_CONFIDENCE"] = (
            "BANK_ONLY_CANDIDATE"
        )

        unmatched_bank.append(
            record
        )

    return (
        pd.DataFrame(matched),
        pd.DataFrame(unmatched_internal),
        pd.DataFrame(unmatched_bank)
    )


# ============================================================
# BNIVA ENGINE
# ============================================================
#
# PENTING:
# - Seluruh fungsi BRIVA di atas dan fast_match() lama di bawah
#   tetap dipertahankan.
# - BNIVA menggunakan engine terpisah agar tidak mengubah logic BRIVA.
# - Rule utama BNIVA:
#       FMSS reff_number 6 digit terakhir == Journal No. Bank
#       KODE_VA sama
#       NOMINAL FMSS == CREDIT Bank
# - Window pencarian bank: D-1 / D / D+1.
# - Fallback hanya VA + nominal yang UNIQUE setelah strong match selesai.
# ============================================================

BNIVA_VA_REGEX = r"(?<!\d)(98876\d{11})(?!\d)"
BNIVA_REFF_REGEX = r"(?i)reff[_\s-]*number\s*=\s*(\d+)"


def extract_bniva_va_series(series):
    """
    Ekstrak VA BNIVA.

    Rule yang sudah tervalidasi pada sampel:
        - prefix 98876
        - panjang 16 digit
    """

    result = (
        series.astype("string")
        .str.extract(
            BNIVA_VA_REGEX,
            expand=False
        )
    )

    return result.where(
        result.notna(),
        None
    )


def classify_bniva_va_series(series):

    result = pd.Series(
        "INVALID VA",
        index=series.index,
        dtype="object"
    )

    mask_bniva = (
        series.astype("string")
        .str.fullmatch(
            r"98876\d{11}",
            na=False
        )
    )

    result.loc[mask_bniva] = "BNIVA"

    return result


def extract_bniva_fmss_journal_series(series):
    """
    Mengambil reff_number dari keterangan FMSS,
    lalu menggunakan 6 digit terakhir sebagai FMSS_JOURNAL.

    Contoh:
        reff_number = 202608182359956872
        FMSS_JOURNAL = 956872
    """

    raw_reff = (
        series.astype("string")
        .str.extract(
            BNIVA_REFF_REGEX,
            expand=False
        )
    )

    journal = (
        raw_reff.astype("string")
        .str[-6:]
    )

    valid_mask = (
        raw_reff.notna()
        & raw_reff.astype("string").str.len().ge(6)
    )

    return journal.where(
        valid_mask,
        None
    )


def normalize_bniva_journal_value(value):
    """
    Normalisasi Journal No. bank menjadi string 6 digit.

    Contoh:
        956872   -> "956872"
        12345    -> "012345"
        956872.0 -> "956872"
    """

    if pd.isna(value):
        return None

    text = str(value).strip()

    if text == "":
        return None

    # Jika dibaca sebagai float, hilangkan .0 di belakang.
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]

    digits = re.sub(
        r"\D",
        "",
        text
    )

    if digits == "":
        return None

    if len(digits) < 6:
        digits = digits.zfill(6)

    elif len(digits) > 6:
        digits = digits[-6:]

    return digits


def normalize_bniva_journal_series(series):

    return series.apply(
        normalize_bniva_journal_value
    )


def parse_bniva_datetime(series):
    """
    Parser khusus format tanggal BNIVA.

    Format sampel:
        17/08/26 23.15.48
    """

    if pd.api.types.is_datetime64_any_dtype(series):

        return pd.to_datetime(
            series,
            errors="coerce"
        )

    text = (
        series.astype("string")
        .str.strip()
    )

    parsed = pd.to_datetime(
        text,
        format="%d/%m/%y %H.%M.%S",
        errors="coerce"
    )

    # Fallback jika export BNI di kemudian hari berubah
    # menjadi format tanggal yang lebih standar.
    mask_fallback = parsed.isna()

    if mask_fallback.any():

        parsed_fallback = pd.to_datetime(
            text[mask_fallback],
            errors="coerce",
            dayfirst=True
        )

        parsed.loc[mask_fallback] = (
            parsed_fallback
        )

    return parsed


def parse_bniva_fmss_datetime(series):
    """
    Parser khusus tanggal FMSS untuk BNIVA.

    Alasan fungsi ini terpisah dari parser bank dan parser umum:
    export FMSS dapat berisi timestamp campuran dalam satu file, misalnya:
        2026-08-19 23:53:00.774091
        2026-08-19 19:46:05

    Pandas dapat menginfer satu format dari baris mayoritas sehingga baris
    tanpa microsecond berubah menjadi NaT. BNIVA pernah mengalami kondisi
    ini pada transaksi retry yang valid.

    Fungsi ini mencoba format dengan microsecond dan tanpa microsecond
    secara eksplisit, lalu memakai fallback hanya untuk nilai yang belum
    berhasil diparse.
    """

    if pd.api.types.is_datetime64_any_dtype(series):

        return pd.to_datetime(
            series,
            errors="coerce"
        )

    text = (
        series.astype("string")
        .str.strip()
    )

    parsed = pd.Series(
        pd.NaT,
        index=series.index,
        dtype="datetime64[ns]"
    )

    # Format standar FMSS dengan microsecond.
    parsed_micro = pd.to_datetime(
        text,
        format="%Y-%m-%d %H:%M:%S.%f",
        errors="coerce"
    )

    parsed = parsed.fillna(
        parsed_micro
    )

    # Format transaksi retry FMSS yang tidak memiliki microsecond.
    parsed_seconds = pd.to_datetime(
        text,
        format="%Y-%m-%d %H:%M:%S",
        errors="coerce"
    )

    parsed = parsed.fillna(
        parsed_seconds
    )

    # Fallback untuk antisipasi export FMSS di kemudian hari berubah.
    mask_fallback = (
        parsed.isna()
        & text.notna()
    )

    if mask_fallback.any():

        try:
            parsed_fallback = pd.to_datetime(
                text[mask_fallback],
                format="mixed",
                errors="coerce"
            )

        except (TypeError, ValueError):
            parsed_fallback = pd.to_datetime(
                text[mask_fallback],
                errors="coerce"
            )

        parsed.loc[mask_fallback] = (
            parsed_fallback
        )

    return parsed


def build_bniva_search_dates(recon_dates):
    """
    Membentuk search window D-1 / D / D+1
    untuk setiap tanggal FMSS yang direkonsiliasi.
    """

    search_dates = set()

    for recon_date in recon_dates:

        date_value = pd.to_datetime(
            recon_date
        ).date()

        search_dates.add(
            date_value - timedelta(days=1)
        )

        search_dates.add(
            date_value
        )

        search_dates.add(
            date_value + timedelta(days=1)
        )

    return search_dates


def prepare_bniva_bank_dataframe(
    uploaded_file,
    recon_dates,
    source_bank="BNIVA"
):
    """
    Load dan normalisasi mutasi BNIVA.

    Search pool:
        tanggal FMSS D-1 / D / D+1

    Hanya credit > 0 yang diproses.

    Transaksi valid neighbor date tetap dipakai sebagai search pool.
    Transaksi tanpa VA pada neighbor date tidak ikut menambah Invalid VA
    rekonsiliasi tanggal target.
    """

    df = read_uploaded_file(
        uploaded_file
    )

    col_credit = find_column(
        df,
        [
            "Credit",
            "CREDIT",
            "credit",
            "KREDIT",
            "kredit",
            "MUTASI_KREDIT",
            "mutasi_kredit"
        ]
    )

    col_desc = find_column(
        df,
        [
            "Description",
            "DESCRIPTION",
            "description",
            "KETERANGAN",
            "keterangan",
            "DESK_TRAN",
            "desk_tran"
        ]
    )

    col_date = find_column(
        df,
        [
            "Post Date",
            "POST DATE",
            "post date",
            "POST_DATE",
            "post_date",
            "TANGGAL",
            "tanggal",
            "TGL_TRAN",
            "tgl_tran"
        ]
    )

    col_journal = find_column(
        df,
        [
            "Journal No.",
            "JOURNAL NO.",
            "Journal No",
            "JOURNAL NO",
            "JOURNAL_NO",
            "journal_no",
            "JOURNAL",
            "journal"
        ]
    )

    df = df.copy()

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    df["_TANGGAL_DT"] = (
        parse_bniva_datetime(
            df[col_date]
        )
    )

    df["_TANGGAL_ONLY_DATE"] = (
        df["_TANGGAL_DT"]
        .dt.date
    )

    # --------------------------------------------------------
    # CREDIT
    # --------------------------------------------------------

    df["_CREDIT_NUM"] = (
        clean_numeric(
            df[col_credit]
        )
    )

    # --------------------------------------------------------
    # SEARCH WINDOW D-1 / D / D+1
    # --------------------------------------------------------

    search_dates = (
        build_bniva_search_dates(
            recon_dates
        )
    )

    target_dates = {
        pd.to_datetime(d).date()
        for d in recon_dates
    }

    df = df[
        df["_TANGGAL_ONLY_DATE"]
        .isin(search_dates)
    ].copy()

    # --------------------------------------------------------
    # HANYA UANG MASUK
    # --------------------------------------------------------

    df = df[
        df["_CREDIT_NUM"] > 0
    ].copy()

    # --------------------------------------------------------
    # VA
    # --------------------------------------------------------

    df["KODE_VA"] = (
        extract_bniva_va_series(
            df[col_desc]
        )
    )

    df["JENIS_VA"] = (
        classify_bniva_va_series(
            df["KODE_VA"]
        )
    )

    # Neighbor date hanya diperlukan sebagai search pool jika VA valid.
    # Invalid VA tetap disimpan jika terjadi pada tanggal target D.
    mask_valid_va = (
        df["KODE_VA"].notna()
    )

    mask_target_date = (
        df["_TANGGAL_ONLY_DATE"]
        .isin(target_dates)
    )

    df = df[
        mask_valid_va
        | mask_target_date
    ].copy()

    # --------------------------------------------------------
    # JOURNAL
    # --------------------------------------------------------

    df["BANK_JOURNAL"] = (
        normalize_bniva_journal_series(
            df[col_journal]
        )
    )

    # --------------------------------------------------------
    # BANK TYPE / SOURCE / DESCRIPTION
    # --------------------------------------------------------

    df["_BANK_TYPE"] = "BNIVA"
    df["SOURCE_BANK"] = source_bank

    df["_DESC_VALUE"] = (
        df[col_desc]
        .astype(str)
    )

    return df


def get_bniva_date_relation(
    int_row,
    bank_row
):
    """
    Label relasi tanggal pasangan BNIVA.
    """

    try:

        fmss_date = pd.to_datetime(
            int_row.get("_TANGGAL_DT")
        ).date()

        bank_date = pd.to_datetime(
            bank_row.get("_TANGGAL_DT")
        ).date()

        delta_days = (
            bank_date - fmss_date
        ).days

    except Exception:

        return "UNKNOWN"

    if delta_days == -1:
        return "H-1 RETRY"

    if delta_days == 0:
        return "SAME DAY"

    if delta_days == 1:
        return "H+1 CUTOFF"

    return f"{delta_days:+d} DAY"


def fast_match_bniva(
    df_int_valid,
    df_bank_valid,
    recon_dates
):
    """
    Matching khusus BNIVA.

    PRIORITAS 1 - STRONG MATCH
        FMSS_JOURNAL == BANK_JOURNAL
        KODE_VA sama
        EXPECTED_BANK == CREDIT bank

    PRIORITAS 2 - FALLBACK
        KODE_VA sama
        EXPECTED_BANK == CREDIT bank
        dan pasangan harus UNIQUE pada remaining rows.

    Matching selalu 1-to-1.

    Bank D-1 dan D+1 digunakan sebagai search pool untuk FMSS D,
    tetapi unmatched neighbor date tidak dihitung sebagai Issue Bank D.
    """

    int_records = (
        df_int_valid
        .to_dict("records")
    )

    bank_records = (
        df_bank_valid
        .to_dict("records")
    )

    matched = []
    unmatched_internal = []
    unmatched_bank = []

    matched_int_indexes = set()
    matched_bank_indexes = set()

    # Jika strong key atau fallback key tidak unik,
    # record diblok dari auto matching agar tidak salah pairing.
    blocked_int_indexes = set()
    blocked_bank_indexes = set()

    # --------------------------------------------------------
    # HELPER ADD MATCH
    # --------------------------------------------------------

    def add_match(
        int_idx,
        bank_idx,
        match_method,
        match_confidence
    ):

        int_row = int_records[int_idx]
        bank_row = bank_records[bank_idx]

        record = int_row.copy()

        record["MATCH_MUTASI_KREDIT"] = (
            bank_row.get(
                "_CREDIT_NUM",
                0
            )
        )

        record["MATCH_DESK_TRAN"] = (
            bank_row.get(
                "_DESC_VALUE",
                ""
            )
        )

        record["SOURCE_BANK"] = (
            bank_row.get(
                "SOURCE_BANK",
                "BNIVA"
            )
        )

        record["BANK_TYPE"] = (
            bank_row.get(
                "_BANK_TYPE",
                "BNIVA"
            )
        )

        record["BANK_JOURNAL"] = (
            bank_row.get(
                "BANK_JOURNAL"
            )
        )

        record["MATCH_BANK_DATE"] = (
            bank_row.get(
                "_TANGGAL_DT"
            )
        )

        record["MATCH_METHOD"] = (
            match_method
        )

        record["MATCH_CONFIDENCE"] = (
            match_confidence
        )

        record["DATE_RELATION"] = (
            get_bniva_date_relation(
                int_row,
                bank_row
            )
        )

        # Tetap menggunakan status MATCHED agar
        # engine dashboard umum tidak perlu berubah.
        record["STATUS_MATCH"] = (
            "MATCHED"
        )

        matched.append(
            record
        )

        matched_int_indexes.add(
            int_idx
        )

        matched_bank_indexes.add(
            bank_idx
        )

    # ========================================================
    # LEVEL A - STRONG MATCH
    # Journal + VA + Nominal
    # ========================================================

    int_strong_index = defaultdict(list)
    bank_strong_index = defaultdict(list)

    for int_idx, int_row in enumerate(
        int_records
    ):

        journal = int_row.get(
            "FMSS_JOURNAL"
        )

        if (
            journal is None
            or str(journal).strip() == ""
        ):
            continue

        key = (
            str(journal),
            str(int_row.get("KODE_VA")),
            float(
                int_row.get(
                    "EXPECTED_BANK",
                    0
                )
            )
        )

        int_strong_index[key].append(
            int_idx
        )

    for bank_idx, bank_row in enumerate(
        bank_records
    ):

        journal = bank_row.get(
            "BANK_JOURNAL"
        )

        if (
            journal is None
            or str(journal).strip() == ""
        ):
            continue

        key = (
            str(journal),
            str(bank_row.get("KODE_VA")),
            float(
                bank_row.get(
                    "_CREDIT_NUM",
                    0
                )
            )
        )

        bank_strong_index[key].append(
            bank_idx
        )

    for key, int_indexes in (
        int_strong_index.items()
    ):

        bank_indexes = (
            bank_strong_index.get(
                key,
                []
            )
        )

        # Auto match hanya jika kedua sisi unique.
        if (
            len(int_indexes) == 1
            and len(bank_indexes) == 1
        ):

            add_match(
                int_indexes[0],
                bank_indexes[0],
                "JOURNAL_VA_NOMINAL",
                "STRONG"
            )

        # Exact strong key tetapi duplicate/ambigu.
        elif len(bank_indexes) > 0:

            blocked_int_indexes.update(
                int_indexes
            )

            blocked_bank_indexes.update(
                bank_indexes
            )

    # ========================================================
    # LEVEL B - FALLBACK MATCH
    # VA + Nominal harus UNIQUE pada remaining rows
    # ========================================================

    int_fallback_index = defaultdict(list)
    bank_fallback_index = defaultdict(list)

    for int_idx, int_row in enumerate(
        int_records
    ):

        if int_idx in matched_int_indexes:
            continue

        if int_idx in blocked_int_indexes:
            continue

        key = (
            str(int_row.get("KODE_VA")),
            float(
                int_row.get(
                    "EXPECTED_BANK",
                    0
                )
            )
        )

        int_fallback_index[key].append(
            int_idx
        )

    for bank_idx, bank_row in enumerate(
        bank_records
    ):

        if bank_idx in matched_bank_indexes:
            continue

        if bank_idx in blocked_bank_indexes:
            continue

        key = (
            str(bank_row.get("KODE_VA")),
            float(
                bank_row.get(
                    "_CREDIT_NUM",
                    0
                )
            )
        )

        bank_fallback_index[key].append(
            bank_idx
        )

    for key, int_indexes in (
        int_fallback_index.items()
    ):

        bank_indexes = (
            bank_fallback_index.get(
                key,
                []
            )
        )

        if (
            len(int_indexes) == 1
            and len(bank_indexes) == 1
        ):

            add_match(
                int_indexes[0],
                bank_indexes[0],
                "VA_NOMINAL_UNIQUE",
                "FALLBACK"
            )

        elif len(bank_indexes) > 0:

            blocked_int_indexes.update(
                int_indexes
            )

            blocked_bank_indexes.update(
                bank_indexes
            )

    # ========================================================
    # FMSS YANG BELUM MATCH
    # ========================================================

    for int_idx, int_row in enumerate(
        int_records
    ):

        if int_idx in matched_int_indexes:
            continue

        record = int_row.copy()

        record["MATCH_METHOD"] = ""
        record["MATCH_CONFIDENCE"] = ""
        record["DATE_RELATION"] = ""

        # ----------------------------------------------------
        # AMBIGUOUS
        # ----------------------------------------------------

        if int_idx in blocked_int_indexes:

            record["STATUS_MATCH"] = (
                "AMBIGUOUS_MATCH"
            )

            unmatched_internal.append(
                record
            )

            continue

        # ----------------------------------------------------
        # DIAGNOSTIK JOURNAL
        # ----------------------------------------------------

        fmss_journal = int_row.get(
            "FMSS_JOURNAL"
        )

        same_journal_bank = []

        if (
            fmss_journal is not None
            and str(fmss_journal).strip() != ""
        ):

            for bank_idx, bank_row in enumerate(
                bank_records
            ):

                if bank_idx in matched_bank_indexes:
                    continue

                if (
                    str(
                        bank_row.get(
                            "BANK_JOURNAL"
                        )
                    )
                    == str(fmss_journal)
                ):

                    same_journal_bank.append(
                        bank_row
                    )

        if same_journal_bank:

            same_journal_va = [
                bank_row
                for bank_row
                in same_journal_bank
                if str(
                    bank_row.get(
                        "KODE_VA"
                    )
                ) == str(
                    int_row.get(
                        "KODE_VA"
                    )
                )
            ]

            same_journal_nominal = [
                bank_row
                for bank_row
                in same_journal_bank
                if float(
                    bank_row.get(
                        "_CREDIT_NUM",
                        0
                    )
                ) == float(
                    int_row.get(
                        "EXPECTED_BANK",
                        0
                    )
                )
            ]

            if same_journal_va:

                record["STATUS_MATCH"] = (
                    "NOMINAL_MISMATCH"
                )

            elif same_journal_nominal:

                record["STATUS_MATCH"] = (
                    "VA_MISMATCH"
                )

            else:

                record["STATUS_MATCH"] = (
                    "JOURNAL_CONFLICT"
                )

        else:

            record["STATUS_MATCH"] = (
                "FMSS_ONLY"
            )

        unmatched_internal.append(
            record
        )

    # ========================================================
    # BANK YANG BELUM MATCH
    # ========================================================
    #
    # Neighbor date D-1 / D+1 hanya berfungsi sebagai search pool.
    # Unmatched neighbor date TIDAK dihitung sebagai Issue Bank D.
    # ========================================================

    target_dates = {
        pd.to_datetime(d).date()
        for d in recon_dates
    }

    fmss_available_dates = set(
        pd.to_datetime(
            df_int_valid["_TANGGAL_DT"],
            errors="coerce"
        )
        .dropna()
        .dt.date
    )

    for bank_idx, bank_row in enumerate(
        bank_records
    ):

        if bank_idx in matched_bank_indexes:
            continue

        bank_datetime = pd.to_datetime(
            bank_row.get(
                "_TANGGAL_DT"
            ),
            errors="coerce"
        )

        if pd.isna(bank_datetime):
            continue

        bank_date = (
            bank_datetime.date()
        )

        # Jangan hitung unmatched D-1 / D+1 sebagai issue tanggal D.
        if bank_date not in target_dates:
            continue

        record = bank_row.copy()

        if bank_idx in blocked_bank_indexes:

            record["STATUS_MATCH"] = (
                "AMBIGUOUS_MATCH - BNIVA"
            )

        else:

            required_fmss_dates = {
                bank_date - timedelta(days=1),
                bank_date,
                bank_date + timedelta(days=1)
            }

            # BANK_ONLY hanya boleh disebut pasti jika FMSS neighbor
            # untuk D-1 / D / D+1 memang tersedia di file FMSS.
            if required_fmss_dates.issubset(
                fmss_available_dates
            ):

                record["STATUS_MATCH"] = (
                    "BANK_ONLY - BNIVA"
                )

            else:

                record["STATUS_MATCH"] = (
                    "BANK_UNVERIFIED - BNIVA"
                )

        unmatched_bank.append(
            record
        )

    # ========================================================
    # DATAFRAME
    # ========================================================

    df_matched = pd.DataFrame(
        matched
    )

    df_selisih_int = pd.DataFrame(
        unmatched_internal
    )

    df_selisih_bnk = pd.DataFrame(
        unmatched_bank
    )

    return (
        df_matched,
        df_selisih_int,
        df_selisih_bnk
    )



# ============================================================
# MANDIRIVA ENGINE - OPTION C / TIERED CONFIDENCE
# ============================================================
#
# PENTING:
# - Engine BRIVA dan BNIVA tidak diubah.
# - MANDIRIVA menggunakan engine terpisah.
# - Prefix yang direkonsiliasi: 888984
# - Fee MANDIRIVA: Rp1.000
# - Search window bank: D-1 / D / D+1
# - Matching Option C:
#       Tahap 1 : VA + nominal UNIQUE -> HIGH CONFIDENCE
#       Tahap 2 : duplicate -> chronological / time resolution 1-to-1
#       Jika waktu benar-benar tidak bisa membedakan -> AMBIGUOUS
# - Unmatched bank D-1 / D+1 tidak dihitung sebagai Issue Bank D.
# - Unmatched bank D hanya boleh menjadi BANK_ONLY jika FMSS D-1/D/D+1
#   tersedia sehingga benar-benar dapat diverifikasi.
# ============================================================

MANDIRIVA_PREFIX = "888984"
MANDIRIVA_FEE = 1000
MANDIRIVA_VA_REGEX = r"(888984\d+)"


def extract_mandiriva_va_series(series):
    """
    Ekstrak VA MANDIRIVA berdasarkan prefix 888984.

    Panjang VA tidak di-hardcode karena pada sampel valid
    ditemukan panjang yang bervariasi.
    """

    result = (
        series.astype("string")
        .str.extract(
            MANDIRIVA_VA_REGEX,
            expand=False
        )
    )

    return result.where(
        result.notna(),
        None
    )


def classify_mandiriva_va_series(series):

    result = pd.Series(
        "INVALID VA",
        index=series.index,
        dtype="object"
    )

    mask_mandiriva = (
        series.astype("string")
        .str.startswith(
            MANDIRIVA_PREFIX,
            na=False
        )
    )

    result.loc[mask_mandiriva] = "MANDIRIVA"

    return result


def parse_mandiriva_datetime(series):
    """
    Parser tanggal mutasi Mandiri.

    Format sampel:
        18 August 2026 01:29:26
    """

    if pd.api.types.is_datetime64_any_dtype(series):

        return pd.to_datetime(
            series,
            errors="coerce"
        )

    text = (
        series.astype("string")
        .str.strip()
    )

    parsed = pd.to_datetime(
        text,
        format="%d %B %Y %H:%M:%S",
        errors="coerce"
    )

    # Fallback apabila format export berubah.
    mask_fallback = parsed.isna()

    if mask_fallback.any():

        parsed_fallback = pd.to_datetime(
            text[mask_fallback],
            errors="coerce",
            dayfirst=True
        )

        parsed.loc[mask_fallback] = (
            parsed_fallback
        )

    return parsed


def build_mandiriva_search_dates(recon_dates):
    """
    Window pencarian MANDIRIVA: D-1 / D / D+1.
    Weekend/libur tidak mengubah window.
    """

    search_dates = set()

    for recon_date in recon_dates:

        date_value = pd.to_datetime(
            recon_date
        ).date()

        search_dates.add(
            date_value - timedelta(days=1)
        )

        search_dates.add(
            date_value
        )

        search_dates.add(
            date_value + timedelta(days=1)
        )

    return search_dates


def prepare_mandiriva_bank_dataframe(
    uploaded_file,
    recon_dates,
    source_bank="MANDIRIVA"
):
    """
    Load dan normalisasi mutasi MANDIRIVA.

    Hanya transaksi yang:
        - berada pada window D-1 / D / D+1
        - Credit Amount > 0
        - berada dalam scope prefix 888984

    Mutasi produk Mandiri lain di rekening yang sama tidak dianggap
    Invalid VA karena memang bukan scope MANDIRIVA 888984.
    """

    df = read_uploaded_file(
        uploaded_file
    )

    col_credit = find_column(
        df,
        [
            "Credit Amount",
            "CREDIT AMOUNT",
            "credit amount",
            "CreditAmount",
            "CREDIT_AMOUNT",
            "credit_amount",
            "CREDIT",
            "Credit",
            "credit",
            "KREDIT",
            "kredit"
        ]
    )

    col_desc = find_column(
        df,
        [
            "Remarks",
            "REMARKS",
            "remarks",
            "AdditionalDesc",
            "ADDITIONALDESC",
            "additionaldesc",
            "DESCRIPTION",
            "Description",
            "description",
            "KETERANGAN",
            "keterangan"
        ]
    )

    col_date = find_column(
        df,
        [
            "PostDate",
            "POSTDATE",
            "postdate",
            "Post Date",
            "POST DATE",
            "post date",
            "POST_DATE",
            "post_date",
            "TANGGAL",
            "tanggal"
        ]
    )

    df = df.copy()

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    df["_TANGGAL_DT"] = (
        parse_mandiriva_datetime(
            df[col_date]
        )
    )

    df["_TANGGAL_ONLY_DATE"] = (
        df["_TANGGAL_DT"]
        .dt.date
    )

    # --------------------------------------------------------
    # CREDIT
    # --------------------------------------------------------

    df["_CREDIT_NUM"] = (
        clean_numeric(
            df[col_credit]
        )
    )

    # --------------------------------------------------------
    # SEARCH WINDOW
    # --------------------------------------------------------

    search_dates = (
        build_mandiriva_search_dates(
            recon_dates
        )
    )

    df = df[
        df["_TANGGAL_ONLY_DATE"]
        .isin(search_dates)
    ].copy()

    # --------------------------------------------------------
    # HANYA UANG MASUK
    # --------------------------------------------------------

    df = df[
        df["_CREDIT_NUM"] > 0
    ].copy()

    # --------------------------------------------------------
    # HANYA SCOPE PREFIX 888984
    # --------------------------------------------------------

    desc_text = (
        df[col_desc]
        .astype("string")
        .fillna("")
    )

    scope_mask = (
        desc_text.str.contains(
            MANDIRIVA_PREFIX,
            regex=False,
            na=False
        )
    )

    df = df[
        scope_mask
    ].copy()

    # --------------------------------------------------------
    # VA
    # --------------------------------------------------------

    df["KODE_VA"] = (
        extract_mandiriva_va_series(
            df[col_desc]
        )
    )

    df["JENIS_VA"] = (
        classify_mandiriva_va_series(
            df["KODE_VA"]
        )
    )

    # --------------------------------------------------------
    # BANK TYPE / SOURCE / DESCRIPTION
    # --------------------------------------------------------

    df["_BANK_TYPE"] = "MANDIRIVA"
    df["SOURCE_BANK"] = source_bank

    df["_DESC_VALUE"] = (
        df[col_desc]
        .astype(str)
    )

    return df


def get_mandiriva_date_relation(
    int_row,
    bank_row
):
    """
    Label relasi tanggal pasangan MANDIRIVA.
    """

    try:

        fmss_date = pd.to_datetime(
            int_row.get("_TANGGAL_DT")
        ).date()

        bank_date = pd.to_datetime(
            bank_row.get("_TANGGAL_DT")
        ).date()

        delta_days = (
            bank_date - fmss_date
        ).days

    except Exception:

        return "UNKNOWN"

    if delta_days == -1:
        return "H-1 RETRY"

    if delta_days == 0:
        return "SAME DAY"

    if delta_days == 1:
        return "H+1 CUTOFF"

    return "OUTSIDE WINDOW"


def mandiriva_time_difference_seconds(
    int_row,
    bank_row
):

    fmss_dt = pd.to_datetime(
        int_row.get("_TANGGAL_DT"),
        errors="coerce"
    )

    bank_dt = pd.to_datetime(
        bank_row.get("_TANGGAL_DT"),
        errors="coerce"
    )

    if pd.isna(fmss_dt) or pd.isna(bank_dt):
        return None

    return (
        bank_dt - fmss_dt
    ).total_seconds()


def _mandiriva_alignment(
    fmss_items,
    bank_items
):
    """
    Chronological sequence alignment untuk duplicate key.

    Objective:
        1. Maksimalkan jumlah pasangan.
        2. Dari jumlah pasangan maksimum, minimalkan total selisih waktu absolut.
        3. Urutan waktu dipertahankan agar transaksi tidak saling silang.

    Tidak membutuhkan scipy / dependency tambahan.
    """

    m = len(fmss_items)
    n = len(bank_items)

    if m == 0 or n == 0:
        return []

    # dp[i][j] = (jumlah_match, total_cost_seconds, path)
    # path berisi tuple posisi (fmss_pos, bank_pos).
    dp = [
        [None for _ in range(n + 1)]
        for _ in range(m + 1)
    ]

    dp[0][0] = (0, 0.0, [])

    for i in range(m + 1):

        for j in range(n + 1):

            current = dp[i][j]

            if current is None:
                continue

            current_matches, current_cost, current_path = current

            # ------------------------------------------------
            # SKIP FMSS
            # ------------------------------------------------

            if i < m:

                candidate = (
                    current_matches,
                    current_cost,
                    current_path
                )

                existing = dp[i + 1][j]

                if (
                    existing is None
                    or candidate[0] > existing[0]
                    or (
                        candidate[0] == existing[0]
                        and candidate[1] < existing[1]
                    )
                ):

                    dp[i + 1][j] = candidate

            # ------------------------------------------------
            # SKIP BANK
            # ------------------------------------------------

            if j < n:

                candidate = (
                    current_matches,
                    current_cost,
                    current_path
                )

                existing = dp[i][j + 1]

                if (
                    existing is None
                    or candidate[0] > existing[0]
                    or (
                        candidate[0] == existing[0]
                        and candidate[1] < existing[1]
                    )
                ):

                    dp[i][j + 1] = candidate

            # ------------------------------------------------
            # MATCH
            # ------------------------------------------------

            if i < m and j < n:

                fmss_dt = fmss_items[i][1]
                bank_dt = bank_items[j][1]

                if (
                    pd.isna(fmss_dt)
                    or pd.isna(bank_dt)
                ):

                    pair_cost = 10 ** 12

                else:

                    pair_cost = abs(
                        (
                            bank_dt
                            - fmss_dt
                        ).total_seconds()
                    )

                candidate = (
                    current_matches + 1,
                    current_cost + pair_cost,
                    current_path + [
                        (i, j)
                    ]
                )

                existing = dp[i + 1][j + 1]

                if (
                    existing is None
                    or candidate[0] > existing[0]
                    or (
                        candidate[0] == existing[0]
                        and candidate[1] < existing[1]
                    )
                ):

                    dp[i + 1][j + 1] = candidate

    result = dp[m][n]

    if result is None:
        return []

    return result[2]


def _mandiriva_group_is_ambiguous(
    fmss_items,
    bank_items,
    alignment
):
    """
    Ambiguous hanya jika timestamp benar-benar tidak memberi pembeda.

    Guardrail dibuat konservatif tanpa threshold waktu bisnis yang di-hardcode:
        - timestamp duplicate identik pada key yang sama; atau
        - satu FMSS mempunyai dua kandidat bank dengan jarak waktu identik.

    Jika tidak terjadi kondisi tersebut, chronological alignment digunakan.
    """

    if not alignment:
        return False

    fmss_times = [
        item[1]
        for item in fmss_items
    ]

    bank_times = [
        item[1]
        for item in bank_items
    ]

    valid_fmss_times = [
        value
        for value in fmss_times
        if not pd.isna(value)
    ]

    valid_bank_times = [
        value
        for value in bank_times
        if not pd.isna(value)
    ]

    if len(valid_fmss_times) != len(set(valid_fmss_times)):
        return True

    if len(valid_bank_times) != len(set(valid_bank_times)):
        return True

    for fmss_pos, bank_pos in alignment:

        fmss_dt = fmss_items[fmss_pos][1]
        selected_bank_dt = bank_items[bank_pos][1]

        if pd.isna(fmss_dt) or pd.isna(selected_bank_dt):
            continue

        selected_distance = abs(
            (
                selected_bank_dt
                - fmss_dt
            ).total_seconds()
        )

        equal_distance_count = 0

        for _, candidate_bank_dt, _ in bank_items:

            if pd.isna(candidate_bank_dt):
                continue

            candidate_distance = abs(
                (
                    candidate_bank_dt
                    - fmss_dt
                ).total_seconds()
            )

            if abs(
                candidate_distance
                - selected_distance
            ) < 0.000001:

                equal_distance_count += 1

        if equal_distance_count > 1:
            return True

    return False


def fast_match_mandiriva(
    df_int_valid,
    df_bank_valid,
    recon_dates
):
    """
    MANDIRIVA Option C - Tiered Confidence Matching.

    Tahap 1:
        VA + EXPECTED_BANK yang hanya muncul 1x di FMSS dan 1x di Bank
        -> MATCHED HIGH CONFIDENCE.

    Tahap 2:
        Duplicate VA + nominal
        -> chronological sequence alignment berdasarkan timestamp.

    Matching selalu 1-to-1.

    EXPECTED_BANK untuk MANDIRIVA sudah dihitung sebagai:
        NOMINAL_ASLI + Rp1.000
    sebelum fungsi ini dipanggil.
    """

    bank_records = (
        df_bank_valid
        .to_dict("records")
    )

    int_records = (
        df_int_valid
        .to_dict("records")
    )

    matched_bank_indexes = set()
    matched_int_indexes = set()

    blocked_bank_indexes = set()
    blocked_int_indexes = set()

    matched = []
    unmatched_internal = []

    # ========================================================
    # BUILD KEY INDEX
    # ========================================================

    fmss_index = defaultdict(list)
    bank_index = defaultdict(list)

    for int_idx, int_row in enumerate(
        int_records
    ):

        key = (
            str(int_row.get("KODE_VA")),
            float(int_row.get("EXPECTED_BANK", 0))
        )

        fmss_index[key].append(
            int_idx
        )

    for bank_idx, bank_row in enumerate(
        bank_records
    ):

        key = (
            str(bank_row.get("KODE_VA")),
            float(bank_row.get("_CREDIT_NUM", 0))
        )

        bank_index[key].append(
            bank_idx
        )

    # ========================================================
    # HELPER: SAVE MATCH
    # ========================================================

    def save_match(
        int_idx,
        bank_idx,
        match_method,
        match_confidence
    ):

        int_row = int_records[
            int_idx
        ]

        bank_row = bank_records[
            bank_idx
        ]

        record = int_row.copy()

        record["MATCH_MUTASI_KREDIT"] = (
            bank_row.get(
                "_CREDIT_NUM",
                0
            )
        )

        record["MATCH_DESK_TRAN"] = (
            bank_row.get(
                "_DESC_VALUE",
                ""
            )
        )

        record["SOURCE_BANK"] = (
            bank_row.get(
                "SOURCE_BANK",
                "MANDIRIVA"
            )
        )

        record["BANK_TYPE"] = (
            bank_row.get(
                "_BANK_TYPE",
                "MANDIRIVA"
            )
        )

        record["MATCH_METHOD"] = (
            match_method
        )

        record["MATCH_CONFIDENCE"] = (
            match_confidence
        )

        record["DATE_RELATION"] = (
            get_mandiriva_date_relation(
                int_row,
                bank_row
            )
        )

        time_difference_seconds = (
            mandiriva_time_difference_seconds(
                int_row,
                bank_row
            )
        )

        record["TIME_DIFFERENCE_SECONDS"] = (
            time_difference_seconds
        )

        if time_difference_seconds is None:

            record["TIME_DIFFERENCE_MINUTES"] = None

        else:

            record["TIME_DIFFERENCE_MINUTES"] = (
                time_difference_seconds / 60
            )

        record["BANK_MATCH_DATETIME"] = (
            bank_row.get(
                "_TANGGAL_DT"
            )
        )

        record["STATUS_MATCH"] = "MATCHED"

        matched.append(
            record
        )

        matched_int_indexes.add(
            int_idx
        )

        matched_bank_indexes.add(
            bank_idx
        )

    # ========================================================
    # TAHAP 1 - UNIQUE EXACT
    # ========================================================

    all_keys = set(
        fmss_index.keys()
    )

    for key in all_keys:

        fmss_candidates = (
            fmss_index.get(
                key,
                []
            )
        )

        bank_candidates = (
            bank_index.get(
                key,
                []
            )
        )

        if (
            len(fmss_candidates) == 1
            and len(bank_candidates) == 1
        ):

            save_match(
                fmss_candidates[0],
                bank_candidates[0],
                "VA_NOMINAL_UNIQUE",
                "HIGH"
            )

    # ========================================================
    # TAHAP 2 - DUPLICATE / TIME RESOLUTION
    # ========================================================

    remaining_keys = set()

    for int_idx, int_row in enumerate(
        int_records
    ):

        if int_idx in matched_int_indexes:
            continue

        key = (
            str(int_row.get("KODE_VA")),
            float(int_row.get("EXPECTED_BANK", 0))
        )

        remaining_keys.add(
            key
        )

    for key in remaining_keys:

        fmss_candidates = [
            idx
            for idx in fmss_index.get(
                key,
                []
            )
            if idx not in matched_int_indexes
        ]

        bank_candidates = [
            idx
            for idx in bank_index.get(
                key,
                []
            )
            if idx not in matched_bank_indexes
        ]

        if not fmss_candidates:
            continue

        if not bank_candidates:
            continue

        fmss_items = []

        for int_idx in fmss_candidates:

            int_dt = pd.to_datetime(
                int_records[int_idx].get(
                    "_TANGGAL_DT"
                ),
                errors="coerce"
            )

            fmss_items.append(
                (
                    int_idx,
                    int_dt,
                    int_records[int_idx]
                )
            )

        bank_items = []

        for bank_idx in bank_candidates:

            bank_dt = pd.to_datetime(
                bank_records[bank_idx].get(
                    "_TANGGAL_DT"
                ),
                errors="coerce"
            )

            bank_items.append(
                (
                    bank_idx,
                    bank_dt,
                    bank_records[bank_idx]
                )
            )

        fmss_items.sort(
            key=lambda item: (
                pd.Timestamp.max
                if pd.isna(item[1])
                else item[1],
                item[0]
            )
        )

        bank_items.sort(
            key=lambda item: (
                pd.Timestamp.max
                if pd.isna(item[1])
                else item[1],
                item[0]
            )
        )

        alignment = (
            _mandiriva_alignment(
                fmss_items,
                bank_items
            )
        )

        if not alignment:
            continue

        is_ambiguous = (
            _mandiriva_group_is_ambiguous(
                fmss_items,
                bank_items,
                alignment
            )
        )

        if is_ambiguous:

            for fmss_pos, bank_pos in alignment:

                int_idx = fmss_items[
                    fmss_pos
                ][0]

                bank_idx = bank_items[
                    bank_pos
                ][0]

                blocked_int_indexes.add(
                    int_idx
                )

                blocked_bank_indexes.add(
                    bank_idx
                )

            continue

        for fmss_pos, bank_pos in alignment:

            int_idx = fmss_items[
                fmss_pos
            ][0]

            bank_idx = bank_items[
                bank_pos
            ][0]

            save_match(
                int_idx,
                bank_idx,
                "TIME_RESOLVED",
                "HIGH"
            )

    # ========================================================
    # FMSS YANG BELUM MATCH
    # ========================================================

    for int_idx, int_row in enumerate(
        int_records
    ):

        if int_idx in matched_int_indexes:
            continue

        record = int_row.copy()

        if int_idx in blocked_int_indexes:

            record["STATUS_MATCH"] = (
                "AMBIGUOUS_MATCH - MANDIRIVA"
            )

            record["MATCH_METHOD"] = (
                "TIME_AMBIGUOUS"
            )

            record["MATCH_CONFIDENCE"] = (
                "LOW"
            )

        else:

            record["STATUS_MATCH"] = (
                "FMSS_ONLY"
            )

            record["MATCH_METHOD"] = (
                "NO_MATCH"
            )

            record["MATCH_CONFIDENCE"] = (
                "NONE"
            )

        unmatched_internal.append(
            record
        )

    # ========================================================
    # ISSUE BANK - MANDIRIVA
    # ========================================================
    #
    # Tujuan rule ini:
    # 1. Tetap TIDAK menghitung unmatched Bank D-1 / D+1 sebagai
    #    Issue Bank untuk periode FMSS D karena keduanya hanya search pool.
    # 2. Unmatched Bank pada tanggal target D harus dapat muncul di
    #    dashboard sebagai kandidat BANK_ONLY agar kasus uang masuk bank
    #    tetapi belum tercatat di FMSS tidak tersembunyi.
    # 3. Tetap melindungi batch carry-over/cutoff H-1 Mandiri yang secara
    #    historis masuk ke Bank D pada dini hari. Tanpa FMSS D-1, transaksi
    #    dini hari tersebut belum aman disebut BANK_ONLY.
    #
    # Guard operasional:
    # - Bank D pukul 00:00:00 s.d. sebelum 03:00:00 yang belum match
    #   dianggap BANK_UNVERIFIED_NEIGHBOR dan TIDAK masuk Issue Bank.
    # - Bank D mulai 03:00:00 yang belum match dimasukkan sebagai
    #   BANK_ONLY_CANDIDATE - MANDIRIVA.
    #
    # Guard 03:00 dipakai untuk menahan false positive batch cutoff H-1
    # (contoh historis batch sekitar 01:32-02:xx), tanpa mengubah engine
    # matching FMSS maupun tampilan dashboard.
    # ========================================================

    unmatched_bank = []

    target_dates = {
        pd.to_datetime(d).date()
        for d in recon_dates
    }

    mandiriva_neighbor_guard_hour = 3

    for bank_idx, bank_row in enumerate(
        bank_records
    ):

        if bank_idx in matched_bank_indexes:
            continue

        if bank_idx in blocked_bank_indexes:
            continue

        bank_datetime = pd.to_datetime(
            bank_row.get(
                "_TANGGAL_DT"
            ),
            errors="coerce"
        )

        if pd.isna(bank_datetime):
            continue

        bank_date = (
            bank_datetime.date()
        )

        # D-1 / D+1 hanya search pool untuk FMSS tanggal target.
        if bank_date not in target_dates:
            continue

        # ----------------------------------------------------
        # PROTEKSI CARRY-OVER H-1 DINI HARI
        # ----------------------------------------------------
        # Tanpa file FMSS D-1, unmatched Bank D sebelum pukul 03:00
        # berpotensi besar merupakan settlement/cutoff transaksi D-1.
        # Jangan naikkan menjadi Issue Bank dashboard.
        # ----------------------------------------------------

        if bank_datetime.hour < mandiriva_neighbor_guard_hour:
            continue

        # ----------------------------------------------------
        # BANK ONLY CANDIDATE
        # ----------------------------------------------------
        # Uang sudah benar-benar masuk di Bank tanggal D, tetapi sesudah
        # seluruh proses matching 1-to-1 tidak ada FMSS D yang memakai
        # record ini. Masukkan ke Issue Bank agar gangguan internal dapat
        # terdeteksi oleh dashboard.
        # ----------------------------------------------------

        record = bank_row.copy()

        record["STATUS_MATCH"] = (
            "BANK_ONLY_CANDIDATE - MANDIRIVA"
        )

        record["MATCH_METHOD"] = (
            "NO_FMSS_MATCH_TARGET_DATE"
        )

        record["MATCH_CONFIDENCE"] = (
            "MEDIUM"
        )

        bank_credit = float(
            record.get(
                "_CREDIT_NUM",
                0
            )
            or 0
        )

        # Estimasi nominal FMSS hanya untuk kebutuhan audit/export.
        # Dashboard tetap memakai nominal uang yang benar-benar masuk bank.
        if bank_credit > MANDIRIVA_FEE:
            record["EXPECTED_FMSS_NOMINAL"] = (
                bank_credit - MANDIRIVA_FEE
            )
            record["BANK_ONLY_NOTE"] = (
                "CREDIT_GT_FEE"
            )
        else:
            record["EXPECTED_FMSS_NOMINAL"] = None
            record["BANK_ONLY_NOTE"] = (
                "CREDIT_LE_FEE_REVIEW"
            )

        unmatched_bank.append(
            record
        )

    # ========================================================
    # DATAFRAME
    # ========================================================

    df_matched = pd.DataFrame(
        matched
    )

    df_selisih_int = pd.DataFrame(
        unmatched_internal
    )

    df_selisih_bnk = pd.DataFrame(
        unmatched_bank
    )

    return (
        df_matched,
        df_selisih_int,
        df_selisih_bnk
    )

# ============================================================
# FAST BANK FILE PROCESSOR
# ============================================================

def prepare_bank_dataframe(
    uploaded_file,
    recon_dates,
    source_bank
):
    """
    Load dan normalisasi file bank.
    Logic matching tidak diubah.
    """

    df = read_uploaded_file(
        uploaded_file
    )

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
            "tanggal"
        ]
    )

    df = df.copy()

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    df["_TANGGAL_DT"] = parse_datetime(
        df[col_date]
    )

    # --------------------------------------------------------
    # CREDIT
    # --------------------------------------------------------

    df["_CREDIT_NUM"] = clean_numeric(
        df[col_credit]
    )

    # --------------------------------------------------------
    # FILTER TANGGAL
    # --------------------------------------------------------

    # Menggunakan normalized datetime agar lebih cepat
    recon_datetime = pd.to_datetime(
        recon_dates
    )

    df["_TANGGAL_ONLY"] = (
        df["_TANGGAL_DT"]
        .dt.normalize()
    )

    df = df[
        df["_TANGGAL_ONLY"].isin(
            recon_datetime
        )
    ].copy()

    # --------------------------------------------------------
    # HANYA UANG MASUK
    # --------------------------------------------------------

    df = df[
        df["_CREDIT_NUM"] > 0
    ].copy()

    # --------------------------------------------------------
    # BANK TYPE
    # --------------------------------------------------------

    df["_BANK_TYPE"] = (
        df[col_desc]
        .astype("string")
        .fillna("")
        .str.upper()
    )

    # Tetap menggunakan klasifikasi yang sama
    df["_BANK_TYPE"] = (
        df[col_desc]
        .apply(classify_bank_transaction)
    )

    # --------------------------------------------------------
    # VA
    # --------------------------------------------------------

    df["KODE_VA"] = extract_va_series(
        df[col_desc]
    )

    df["JENIS_VA"] = classify_va_series(
        df["KODE_VA"]
    )

    # --------------------------------------------------------
    # SOURCE
    # --------------------------------------------------------

    df["SOURCE_BANK"] = source_bank

    # --------------------------------------------------------
    # SIMPAN DESCRIPTION
    # agar tidak perlu mencari nama kolom lagi
    # ketika matching
    # --------------------------------------------------------

    df["_DESC_VALUE"] = (
        df[col_desc]
        .astype(str)
    )

    return df


# ============================================================
# FAST MATCHING ENGINE
# ============================================================

def fast_match(
    df_int_valid,
    df_bank_valid
):
    """
    Matching 1-to-1 berbasis dictionary.

    LOGIC SAMA:
        KODE_VA harus sama
        EXPECTED_BANK harus sama dengan CREDIT BANK

    Perbedaan:
        Tidak lagi melakukan nested loop.
    """

    # --------------------------------------------------------
    # BANK RECORDS
    # --------------------------------------------------------

    bank_records = (
        df_bank_valid
        .to_dict("records")
    )

    # --------------------------------------------------------
    # INDEX BANK
    #
    # key:
    #   (KODE_VA, CREDIT)
    #
    # value:
    #   queue index bank
    #
    # deque dipakai agar duplicate transaction
    # tetap diproses 1-to-1 sesuai urutan.
    # --------------------------------------------------------

    bank_index = defaultdict(
        deque
    )

    for idx, bank_row in enumerate(
        bank_records
    ):

        key = (
            str(bank_row["KODE_VA"]),
            float(bank_row["_CREDIT_NUM"])
        )

        bank_index[key].append(
            idx
        )

    # --------------------------------------------------------
    # TRACK BANK YANG SUDAH MATCH
    # --------------------------------------------------------

    matched_bank_indexes = set()

    matched = []
    unmatched_internal = []

    # --------------------------------------------------------
    # FMSS RECORDS
    # --------------------------------------------------------

    int_records = (
        df_int_valid
        .to_dict("records")
    )

    # --------------------------------------------------------
    # MATCH
    # --------------------------------------------------------

    for int_row in int_records:

        key = (
            str(int_row["KODE_VA"]),
            float(int_row["EXPECTED_BANK"])
        )

        queue = bank_index.get(
            key
        )

        # ----------------------------------------------------
        # MATCH FOUND
        # ----------------------------------------------------

        if queue:

            bank_idx = queue.popleft()

            bank_row = (
                bank_records[
                    bank_idx
                ]
            )

            matched_bank_indexes.add(
                bank_idx
            )

            record = int_row.copy()

            record["MATCH_MUTASI_KREDIT"] = (
                bank_row["_CREDIT_NUM"]
            )

            record["MATCH_DESK_TRAN"] = (
                bank_row.get(
                    "_DESC_VALUE",
                    ""
                )
            )

            record["SOURCE_BANK"] = (
                bank_row.get(
                    "SOURCE_BANK",
                    ""
                )
            )

            record["BANK_TYPE"] = (
                bank_row.get(
                    "_BANK_TYPE",
                    ""
                )
            )

            record["STATUS_MATCH"] = (
                "MATCHED"
            )

            matched.append(
                record
            )

        # ----------------------------------------------------
        # FMSS ONLY
        # ----------------------------------------------------

        else:

            record = int_row.copy()

            record["STATUS_MATCH"] = (
                "FMSS_ONLY"
            )

            unmatched_internal.append(
                record
            )

    # --------------------------------------------------------
    # BANK ONLY
    #
    # Tetap berdasarkan urutan asli file bank.
    # --------------------------------------------------------

    unmatched_bank = []

    for idx, bank_row in enumerate(
        bank_records
    ):

        if idx in matched_bank_indexes:
            continue

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

    # --------------------------------------------------------
    # DATAFRAME
    # --------------------------------------------------------

    df_matched = pd.DataFrame(
        matched
    )

    df_selisih_int = pd.DataFrame(
        unmatched_internal
    )

    df_selisih_bnk = pd.DataFrame(
        unmatched_bank
    )

    return (
        df_matched,
        df_selisih_int,
        df_selisih_bnk
    )


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


# ============================================================
# RESET RESULT JIKA BANK BERUBAH
# ============================================================

if (
    pilihan_bank
    != st.session_state.pilihan_bank_terakhir
):

    st.session_state.sudah_diproses = False

    st.session_state.df_matched = (
        pd.DataFrame()
    )

    st.session_state.df_selisih_int = (
        pd.DataFrame()
    )

    st.session_state.df_selisih_bnk = (
        pd.DataFrame()
    )

    st.session_state.df_invalid_int = (
        pd.DataFrame()
    )

    st.session_state.df_invalid_bnk = (
        pd.DataFrame()
    )

    st.session_state.recon_dates = []

    st.session_state.summary = {}

    st.session_state.pilihan_bank_terakhir = (
        pilihan_bank
    )


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

        st.markdown(
            "### 🏦 BRIVA Fastpay — 57888"
        )

        file_bnk_57888 = st.file_uploader(
            "Upload mutasi BRIVA 57888",
            type=["csv", "xlsx"],
            key="briva_57888"
        )

    with col3:

        st.markdown(
            "### 🏦 BRIVA Rajabiller — 57708"
        )

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

        st.markdown(
            f"### 🏦 Mutasi {pilihan_bank}"
        )

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
# BUTTON PROCESS
# ============================================================

can_process = False

if pilihan_bank == "BRIVA":

    if (
        file_int
        and file_bnk_57888
        and file_bnk_57708
    ):
        can_process = True

else:

    if (
        file_int
        and file_bnk_general
    ):
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

                # =================================================
                # LOAD FMSS
                # =================================================

                df_int = read_uploaded_file(
                    file_int
                )

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

                # =================================================
                # FILTER FMSS SUKSES
                # =================================================

                df_int = df_int.copy()

                df_int["_STATUS_CLEAN"] = (
                    df_int[col_status]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df_int_sukses = df_int[
                    df_int["_STATUS_CLEAN"]
                    == "SUKSES"
                ].copy()

                # =================================================
                # TANGGAL REKONSILIASI
                # =================================================

                if pilihan_bank == "BRIVA":

                    df_int_sukses["_TANGGAL_DT"] = (
                        parse_briva_datetime(
                            df_int_sukses[
                                col_tanggal_int
                            ]
                        )
                    )

                elif pilihan_bank == "BNIVA":

                    df_int_sukses["_TANGGAL_DT"] = (
                        parse_bniva_fmss_datetime(
                            df_int_sukses[
                                col_tanggal_int
                            ]
                        )
                    )

                else:

                    df_int_sukses["_TANGGAL_DT"] = (
                        parse_datetime(
                            df_int_sukses[
                                col_tanggal_int
                            ]
                        )
                    )

                df_int_sukses = (
                    df_int_sukses[
                        df_int_sukses[
                            "_TANGGAL_DT"
                        ].notna()
                    ].copy()
                )

                if df_int_sukses.empty:

                    raise ValueError(
                        "Tidak ada transaksi FMSS SUKSES "
                        "dengan tanggal yang valid."
                    )

                recon_dates = sorted(
                    df_int_sukses[
                        "_TANGGAL_DT"
                    ]
                    .dt.date
                    .dropna()
                    .unique()
                )

                st.session_state.recon_dates = (
                    recon_dates
                )

                # =================================================
                # EXTRACT VA FMSS - VECTORIZED
                # =================================================

                if pilihan_bank == "BNIVA":

                    df_int_sukses["KODE_VA"] = (
                        extract_bniva_va_series(
                            df_int_sukses[
                                col_keterangan_int
                            ]
                        )
                    )

                    df_int_sukses["JENIS_VA"] = (
                        classify_bniva_va_series(
                            df_int_sukses[
                                "KODE_VA"
                            ]
                        )
                    )

                    df_int_sukses["FMSS_JOURNAL"] = (
                        extract_bniva_fmss_journal_series(
                            df_int_sukses[
                                col_keterangan_int
                            ]
                        )
                    )

                elif pilihan_bank == "MANDIRIVA":

                    df_int_sukses["KODE_VA"] = (
                        extract_mandiriva_va_series(
                            df_int_sukses[
                                col_keterangan_int
                            ]
                        )
                    )

                    df_int_sukses["JENIS_VA"] = (
                        classify_mandiriva_va_series(
                            df_int_sukses[
                                "KODE_VA"
                            ]
                        )
                    )

                else:

                    df_int_sukses["KODE_VA"] = (
                        extract_va_series(
                            df_int_sukses[
                                col_keterangan_int
                            ]
                        )
                    )

                    df_int_sukses["JENIS_VA"] = (
                        classify_va_series(
                            df_int_sukses[
                                "KODE_VA"
                            ]
                        )
                    )

                # =================================================
                # INVALID VA FMSS
                # =================================================

                df_invalid_int = (
                    df_int_sukses[
                        df_int_sukses[
                            "KODE_VA"
                        ].isna()
                    ].copy()
                )

                # =================================================
                # FMSS VALID
                # =================================================

                df_int_valid = (
                    df_int_sukses[
                        df_int_sukses[
                            "KODE_VA"
                        ].notna()
                    ].copy()
                )

                # =================================================
                # NOMINAL FMSS
                # =================================================

                df_int_valid[
                    "NOMINAL_ASLI"
                ] = clean_numeric(
                    df_int_valid[
                        col_nominal_int
                    ]
                )

                # =================================================
                # EXPECTED BANK
                # =================================================

                df_int_valid[
                    "EXPECTED_BANK"
                ] = (
                    df_int_valid[
                        "NOMINAL_ASLI"
                    ]
                )

                mask_57888 = (
                    df_int_valid[
                        "JENIS_VA"
                    ]
                    == "BRIVA FASTPAY"
                )

                mask_57708 = (
                    df_int_valid[
                        "JENIS_VA"
                    ]
                    == "BRIVA RAJABILLER"
                )

                df_int_valid.loc[
                    mask_57888,
                    "EXPECTED_BANK"
                ] = (
                    df_int_valid.loc[
                        mask_57888,
                        "NOMINAL_ASLI"
                    ]
                    + fee_57888
                )

                df_int_valid.loc[
                    mask_57708,
                    "EXPECTED_BANK"
                ] = (
                    df_int_valid.loc[
                        mask_57708,
                        "NOMINAL_ASLI"
                    ]
                    + fee_57708
                )

                if pilihan_bank == "MANDIRIVA":

                    df_int_valid[
                        "EXPECTED_BANK"
                    ] = (
                        df_int_valid[
                            "NOMINAL_ASLI"
                        ]
                        + MANDIRIVA_FEE
                    )

                # =================================================
                # BANK PROCESSING
                # =================================================

                bank_sources = []

                if pilihan_bank == "BRIVA":

                    # ---------------------------------------------
                    # BRIVA 57888
                    # ---------------------------------------------

                    df_57888 = prepare_briva_bank_dataframe(
                        file_bnk_57888,
                        recon_dates,
                        "BRIVA FASTPAY 57888"
                    )

                    bank_sources.append(
                        df_57888
                    )

                    # ---------------------------------------------
                    # BRIVA 57708
                    # ---------------------------------------------

                    df_57708 = prepare_briva_bank_dataframe(
                        file_bnk_57708,
                        recon_dates,
                        "BRIVA RAJABILLER 57708"
                    )

                    bank_sources.append(
                        df_57708
                    )

                elif pilihan_bank == "BNIVA":

                    # ---------------------------------------------
                    # BNIVA
                    # Engine khusus, terpisah dari BRIVA.
                    # ---------------------------------------------

                    df_bniva = prepare_bniva_bank_dataframe(
                        file_bnk_general,
                        recon_dates,
                        "BNIVA"
                    )

                    bank_sources.append(
                        df_bniva
                    )

                elif pilihan_bank == "MANDIRIVA":

                    # ---------------------------------------------
                    # MANDIRIVA
                    # Engine khusus Option C, terpisah dari
                    # BRIVA dan BNIVA.
                    # ---------------------------------------------

                    df_mandiriva = (
                        prepare_mandiriva_bank_dataframe(
                            file_bnk_general,
                            recon_dates,
                            "MANDIRIVA"
                        )
                    )

                    bank_sources.append(
                        df_mandiriva
                    )

                else:

                    # ---------------------------------------------
                    # BANK LAIN
                    # ---------------------------------------------

                    df_general = prepare_bank_dataframe(
                        file_bnk_general,
                        recon_dates,
                        pilihan_bank
                    )

                    bank_sources.append(
                        df_general
                    )

                # =================================================
                # COMBINE BANK
                # =================================================

                if bank_sources:

                    df_bank = pd.concat(
                        bank_sources,
                        ignore_index=True
                    )

                else:

                    df_bank = pd.DataFrame()

                # =================================================
                # BANK INVALID VA
                # =================================================

                df_invalid_bnk = (
                    df_bank[
                        df_bank[
                            "KODE_VA"
                        ].isna()
                    ].copy()
                )

                # =================================================
                # BANK VALID
                # =================================================

                df_bank_valid = (
                    df_bank[
                        df_bank[
                            "KODE_VA"
                        ].notna()
                    ].copy()
                )

                # =================================================
                # FAST MATCHING ENGINE
                # =================================================

                if pilihan_bank == "BNIVA":

                    (
                        df_matched,
                        df_selisih_int,
                        df_selisih_bnk
                    ) = fast_match_bniva(
                        df_int_valid,
                        df_bank_valid,
                        recon_dates
                    )

                elif pilihan_bank == "MANDIRIVA":

                    (
                        df_matched,
                        df_selisih_int,
                        df_selisih_bnk
                    ) = fast_match_mandiriva(
                        df_int_valid,
                        df_bank_valid,
                        recon_dates
                    )

                elif pilihan_bank == "BRIVA":

                    (
                        df_matched,
                        df_selisih_int,
                        df_selisih_bnk
                    ) = fast_match_briva(
                        df_int_valid,
                        df_bank_valid,
                        recon_dates
                    )

                else:

                    (
                        df_matched,
                        df_selisih_int,
                        df_selisih_bnk
                    ) = fast_match(
                        df_int_valid,
                        df_bank_valid
                    )

                # =================================================
                # SUMMARY
                # =================================================

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
                    df_matched[
                        "NOMINAL_ASLI"
                    ].sum()
                    if (
                        not df_matched.empty
                        and "NOMINAL_ASLI"
                        in df_matched.columns
                    )
                    else 0
                )

                fmss_only_nominal = (
                    df_selisih_int[
                        "NOMINAL_ASLI"
                    ].sum()
                    if (
                        not df_selisih_int.empty
                        and "NOMINAL_ASLI"
                        in df_selisih_int.columns
                    )
                    else 0
                )

                bank_only_nominal = (
                    df_selisih_bnk[
                        "_CREDIT_NUM"
                    ].sum()
                    if (
                        not df_selisih_bnk.empty
                        and "_CREDIT_NUM"
                        in df_selisih_bnk.columns
                    )
                    else 0
                )

                summary = {
                    "matched_count":
                        matched_count,

                    "fmss_only_count":
                        fmss_only_count,

                    "bank_only_count":
                        bank_only_count,

                    "invalid_int_count":
                        invalid_int_count,

                    "invalid_bnk_count":
                        invalid_bnk_count,

                    "matched_nominal":
                        matched_nominal,

                    "fmss_only_nominal":
                        fmss_only_nominal,

                    "bank_only_nominal":
                        bank_only_nominal
                }

                # =================================================
                # SAVE SESSION
                # =================================================

                st.session_state.df_matched = (
                    df_matched
                )

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

                st.session_state.summary = (
                    summary
                )

                st.session_state.sudah_diproses = (
                    True
                )

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

    df_matched = (
        st.session_state.df_matched
    )

    df_selisih_int = (
        st.session_state.df_selisih_int
    )

    df_selisih_bnk = (
        st.session_state.df_selisih_bnk
    )

    df_invalid_int = (
        st.session_state.df_invalid_int
    )

    df_invalid_bnk = (
        st.session_state.df_invalid_bnk
    )

    summary = (
        st.session_state.summary
    )

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

    st.subheader(
        "💰 Ringkasan Nominal"
    )

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

        st.subheader(
            "🚨 Issue FMSS"
        )

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

            display_int["ISSUE"] = (
                "FMSS_ONLY"
            )

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

        st.subheader(
            "🚨 Issue Bank"
        )

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

            st.markdown(
                "### FMSS Invalid VA"
            )

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

            st.markdown(
                "### Bank Invalid VA"
            )

            if not df_invalid_bnk.empty:

                invalid_cols = []

                for col in [
                    "_TANGGAL_DT",
                    "_BANK_TYPE",
                    "SOURCE_BANK"
                ]:

                    if col in df_invalid_bnk.columns:

                        invalid_cols.append(
                            col
                        )

                if "KODE_VA" in df_invalid_bnk.columns:

                    invalid_cols.append(
                        "KODE_VA"
                    )

                if "_CREDIT_NUM" in df_invalid_bnk.columns:

                    invalid_cols.append(
                        "_CREDIT_NUM"
                    )

                st.dataframe(
                    df_invalid_bnk[
                        invalid_cols
                    ],
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

    st.subheader(
        "📥 Download Laporan"
    )

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

                    summary[
                        "matched_count"
                    ],

                    summary[
                        "fmss_only_count"
                    ],

                    summary[
                        "bank_only_count"
                    ],

                    summary[
                        "invalid_int_count"
                    ],

                    summary[
                        "invalid_bnk_count"
                    ],

                    f"{fmss_match_rate:.4f}%",

                    f"{bank_match_rate:.4f}%",

                    summary[
                        "matched_nominal"
                    ],

                    summary[
                        "fmss_only_nominal"
                    ],

                    summary[
                        "bank_only_nominal"
                    ]
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

                export_matched = (
                    df_matched.copy()
                )

                export_matched = (
                    export_matched.drop(
                        columns=[
                            "_STATUS_CLEAN",
                            "_TANGGAL_DT",
                            "_TANGGAL_ONLY"
                        ],
                        errors="ignore"
                    )
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

                export_fmss = (
                    df_selisih_int.copy()
                )

                export_fmss = (
                    export_fmss.drop(
                        columns=[
                            "_STATUS_CLEAN",
                            "_TANGGAL_DT"
                        ],
                        errors="ignore"
                    )
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

                export_bank = (
                    df_selisih_bnk.copy()
                )

                export_bank = (
                    export_bank.drop(
                        columns=[
                            "_TANGGAL_DT",
                            "_TANGGAL_ONLY"
                        ],
                        errors="ignore"
                    )
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
