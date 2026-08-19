import streamlit as st
import pandas as pd
import re
import io
import traceback
from datetime import datetime

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Rekonsiliasi Bank Fastpay",
    layout="wide"
)

st.title("📊 Rekonsiliasi Bank Fastpay")
st.write(
    "Sistem otomatis mencocokkan data deposit internal (FMSS) "
    "dengan mutasi uang masuk di bank."
)

st.divider()


# ============================================================
# EXTRACTOR REGISTRY
# ============================================================

def extract_va_code(text, prefix: str):
    if pd.isna(text):
        return None

    match = re.search(
        rf'({re.escape(prefix)}\d{{5,15}})',
        str(text)
    )

    return match.group(1) if match else None


def extractor_regex_prefix(df, source_col, prefix, **_):
    return df[source_col].apply(
        lambda x: extract_va_code(x, prefix)
    )


def extractor_dedicated_column(df, source_col, **_):
    s = df[source_col].astype(str).str.strip()

    return s.replace({
        'nan': None,
        'None': None,
        '': None
    })


def extractor_split_delimiter(
    df,
    source_col,
    delimiter="-",
    index=-1,
    **_
):
    def _split(x):
        if pd.isna(x):
            return None

        parts = str(x).split(delimiter)

        try:
            value = parts[index].strip()
            return value if value else None
        except IndexError:
            return None

    return df[source_col].apply(_split)


EXTRACTOR_REGISTRY = {
    "regex_prefix": extractor_regex_prefix,
    "dedicated_column": extractor_dedicated_column,
    "split_delimiter": extractor_split_delimiter,
}


def apply_extractor(df: pd.DataFrame, spec: dict) -> pd.Series:
    if spec["fn"] not in EXTRACTOR_REGISTRY:
        raise ValueError(
            f"Extractor '{spec['fn']}' belum tersedia."
        )

    fn = EXTRACTOR_REGISTRY[spec["fn"]]

    kwargs = {
        k: v
        for k, v in spec.items()
        if k != "fn"
    }

    return fn(df, **kwargs)


# ============================================================
# BANK CONFIG
# ============================================================

BANK_CONFIGS = {

    "BRIVA": {
        "configured": True,

        "internal_extractor": {
            "fn": "regex_prefix",
            "source_col": "keterangan",
            "prefix": "57888"
        },

        "bank_extractor": {
            "fn": "regex_prefix",
            "source_col": "DESK_TRAN",
            "prefix": "57888"
        },

        "bank_credit_col": "MUTASI_KREDIT",

        "fee_adjustment": 1000,

        "internal_date_col": None,
        "bank_date_col": None,
    },

    "BNIVA": {
        "configured": False
    },

    "BCAVA": {
        "configured": False
    },

    "MANDIRIVA": {
        "configured": False
    },

    "BSIVA": {
        "configured": False
    },

    "MuamalatVA": {
        "configured": False
    },
}


INTERNAL_REQUIRED_COLS = [
    "status",
    "keterangan",
    "nominal"
]


# ============================================================
# UTILITIES
# ============================================================

def parse_nominal(series: pd.Series) -> pd.Series:

    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0)

    cleaned = (
        series.astype(str)
        .str.replace(r'Rp\.?', '', regex=True)
        .str.replace(' ', '', regex=False)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
        .str.strip()
    )

    return pd.to_numeric(
        cleaned,
        errors='coerce'
    ).fillna(0)


def validate_columns(
    df: pd.DataFrame,
    required_cols: list,
    source_label: str
):

    missing = [
        c for c in required_cols
        if c not in df.columns
    ]

    if missing:
        raise ValueError(
            f"File **{source_label}** tidak memiliki "
            f"kolom wajib: {', '.join(missing)}.\n\n"
            f"Kolom yang tersedia: "
            f"{', '.join(map(str, df.columns))}"
        )


def read_any(uploaded_file):

    uploaded_file.seek(0)

    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python"
        )

    return pd.read_excel(uploaded_file)


def generate_reconciliation_id():

    return datetime.now().strftime(
        "REC-%Y%m%d-%H%M%S"
    )


# ============================================================
# MATCHING ENGINE
# ============================================================

def tally_1_to_1(
    df_int: pd.DataFrame,
    df_bnk: pd.DataFrame
):

    di = df_int.copy()
    db = df_bnk.copy()

    # Occurrence number memastikan duplicate VA + nominal
    # tetap dicocokkan 1-to-1.
    di["_occ"] = (
        di.groupby(
            ["VA_CODE", "NOMINAL_MATCH"]
        ).cumcount()
    )

    db["_occ"] = (
        db.groupby(
            ["VA_CODE", "NOMINAL_MATCH"]
        ).cumcount()
    )

    di = di.add_prefix("INT__")
    db = db.add_prefix("BNK__")

    di = di.rename(columns={
        "INT__VA_CODE": "VA_CODE",
        "INT__NOMINAL_MATCH": "NOMINAL_MATCH",
        "INT___occ": "_occ"
    })

    db = db.rename(columns={
        "BNK__VA_CODE": "VA_CODE",
        "BNK__NOMINAL_MATCH": "NOMINAL_MATCH",
        "BNK___occ": "_occ"
    })

    merged = di.merge(
        db,
        on=[
            "VA_CODE",
            "NOMINAL_MATCH",
            "_occ"
        ],
        how="outer",
        indicator=True
    )

    matched = merged[
        merged["_merge"] == "both"
    ].copy()

    only_int = merged[
        merged["_merge"] == "left_only"
    ].copy()

    only_bnk = merged[
        merged["_merge"] == "right_only"
    ].copy()

    return matched, only_int, only_bnk


# ============================================================
# CLASSIFY UNMATCHED
# ============================================================

def classify_unmatched(
    only_int: pd.DataFrame,
    only_bnk: pd.DataFrame
):

    only_int = only_int.copy()
    only_bnk = only_bnk.copy()

    bank_remaining_va = set(
        only_bnk["VA_CODE"].dropna()
    )

    int_remaining_va = set(
        only_int["VA_CODE"].dropna()
    )

    if not only_int.empty:

        only_int["ISSUE_TYPE"] = (
            only_int["VA_CODE"]
            .isin(bank_remaining_va)
            .map({
                True: "AMOUNT_MISMATCH",
                False: "FMSS_ONLY"
            })
        )

    if not only_bnk.empty:

        only_bnk["ISSUE_TYPE"] = (
            only_bnk["VA_CODE"]
            .isin(int_remaining_va)
            .map({
                True: "AMOUNT_MISMATCH",
                False: "BANK_ONLY"
            })
        )

    return only_int, only_bnk


# ============================================================
# PROCESS RECONCILIATION
# ============================================================

def process_reconciliation(
    df_int_raw: pd.DataFrame,
    df_bnk_raw: pd.DataFrame,
    config: dict,
    reconciliation_id: str,
    df_carry_over: pd.DataFrame = None
):

    # --------------------------------------------------------
    # VALIDATE INPUT
    # --------------------------------------------------------

    validate_columns(
        df_int_raw,
        INTERNAL_REQUIRED_COLS,
        "FMSS (Internal)"
    )

    bank_required = [
        config["bank_extractor"]["source_col"],
        config["bank_credit_col"]
    ]

    validate_columns(
        df_bnk_raw,
        bank_required,
        "Mutasi Bank"
    )

    # --------------------------------------------------------
    # INTERNAL DATA
    # --------------------------------------------------------

    df_int_success = df_int_raw[
        df_int_raw["status"]
        .astype(str)
        .str.upper()
        == "SUKSES"
    ].copy()

    total_internal_success = len(
        df_int_success
    )

    total_internal_success_amount = (
        parse_nominal(
            df_int_success["nominal"]
        ).sum()
    )

    df_int_success["SUMBER_DATA"] = "BARU"

    # --------------------------------------------------------
    # CARRY OVER
    # --------------------------------------------------------

    carry_count = 0

    if (
        df_carry_over is not None
        and not df_carry_over.empty
    ):

        validate_columns(
            df_carry_over,
            INTERNAL_REQUIRED_COLS,
            "Carry-over FMSS"
        )

        df_carry = df_carry_over.copy()

        df_carry["SUMBER_DATA"] = (
            "CARRY_OVER"
        )

        carry_count = len(df_carry)

        df_int_success = pd.concat(
            [
                df_carry,
                df_int_success
            ],
            ignore_index=True
        )

    # --------------------------------------------------------
    # INTERNAL NORMALIZATION
    # --------------------------------------------------------

    df_int_success["VA_CODE"] = (
        apply_extractor(
            df_int_success,
            config["internal_extractor"]
        )
    )

    df_int_success["NOMINAL_ORIGINAL"] = (
        parse_nominal(
            df_int_success["nominal"]
        )
    )

    df_int_success["NOMINAL_MATCH"] = (
        df_int_success["NOMINAL_ORIGINAL"]
        + config["fee_adjustment"]
    )

    invalid_internal = df_int_success[
        df_int_success["VA_CODE"].isna()
    ].copy()

    valid_internal = df_int_success[
        df_int_success["VA_CODE"].notna()
    ].reset_index(drop=True)

    # --------------------------------------------------------
    # BANK DATA
    # --------------------------------------------------------

    df_bnk = df_bnk_raw.copy()

    df_bnk["NOMINAL_KREDIT_NUM"] = (
        parse_nominal(
            df_bnk[
                config["bank_credit_col"]
            ]
        )
    )

    df_bnk_credit = df_bnk[
        df_bnk["NOMINAL_KREDIT_NUM"] > 0
    ].copy()

    total_bank_credit = len(
        df_bnk_credit
    )

    total_bank_credit_amount = (
        df_bnk_credit[
            "NOMINAL_KREDIT_NUM"
        ].sum()
    )

    df_bnk_credit["VA_CODE"] = (
        apply_extractor(
            df_bnk_credit,
            config["bank_extractor"]
        )
    )

    invalid_bank = df_bnk_credit[
        df_bnk_credit["VA_CODE"].isna()
    ].copy()

    valid_bank = df_bnk_credit[
        df_bnk_credit["VA_CODE"].notna()
    ].reset_index(drop=True)

    valid_bank["NOMINAL_MATCH"] = (
        valid_bank["NOMINAL_KREDIT_NUM"]
    )

    # --------------------------------------------------------
    # DATA QUALITY CHECK
    # --------------------------------------------------------

    if valid_internal.empty:

        raise ValueError(
            "Tidak ada transaksi FMSS SUCCESS "
            "dengan VA valid."
        )

    if valid_bank.empty:

        raise ValueError(
            "Tidak ada mutasi kredit bank "
            "dengan VA valid."
        )

    # --------------------------------------------------------
    # MATCH
    # --------------------------------------------------------

    matched, only_int, only_bnk = (
        tally_1_to_1(
            valid_internal,
            valid_bank
        )
    )

    # --------------------------------------------------------
    # CLASSIFICATION
    # --------------------------------------------------------

    only_int, only_bnk = (
        classify_unmatched(
            only_int,
            only_bnk
        )
    )

    # --------------------------------------------------------
    # MATCH METADATA
    # --------------------------------------------------------

    if not matched.empty:

        matched["MATCH_STATUS"] = "MATCHED"

        matched["MATCH_AMOUNT"] = (
            matched["NOMINAL_MATCH"]
        )

        # Date audit
        int_date_col = config.get(
            "internal_date_col"
        )

        bank_date_col = config.get(
            "bank_date_col"
        )

        if (
            int_date_col
            and bank_date_col
        ):

            int_key = (
                f"INT__{int_date_col}"
            )

            bank_key = (
                f"BNK__{bank_date_col}"
            )

            if (
                int_key in matched.columns
                and bank_key in matched.columns
            ):

                d_int = pd.to_datetime(
                    matched[int_key],
                    errors="coerce"
                )

                d_bank = pd.to_datetime(
                    matched[bank_key],
                    errors="coerce"
                )

                matched[
                    "SELISIH_HARI"
                ] = (
                    d_bank - d_int
                ).dt.days

                matched[
                    "MATCH_TIMING"
                ] = matched[
                    "SELISIH_HARI"
                ].apply(
                    lambda x:
                        "UNKNOWN"
                        if pd.isna(x)
                        else (
                            "SAME_DAY"
                            if x == 0
                            else (
                                "H+1"
                                if x == 1
                                else (
                                    "H+2+"
                                    if x > 1
                                    else "BANK_BEFORE"
                                )
                            )
                        )
                )

            else:

                matched[
                    "MATCH_TIMING"
                ] = "UNKNOWN"

        else:

            matched[
                "MATCH_TIMING"
            ] = "UNKNOWN"

        matched[
            "RECONCILIATION_ID"
        ] = reconciliation_id

    # --------------------------------------------------------
    # ISSUE METADATA
    # --------------------------------------------------------

    if not only_int.empty:

        only_int[
            "RECONCILIATION_ID"
        ] = reconciliation_id

    if not only_bnk.empty:

        only_bnk[
            "RECONCILIATION_ID"
        ] = reconciliation_id

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    matched_count = len(matched)

    matched_amount = (
        matched["NOMINAL_MATCH"].sum()
        if not matched.empty
        else 0
    )

    issue_fmss_amount = (
        only_int["NOMINAL_MATCH"].sum()
        if not only_int.empty
        else 0
    )

    issue_bank_amount = (
        only_bnk["NOMINAL_MATCH"].sum()
        if not only_bnk.empty
        else 0
    )

    eligible_internal_count = (
        len(valid_internal)
    )

    eligible_bank_count = (
        len(valid_bank)
    )

    transaction_match_rate = (
        matched_count
        / eligible_internal_count
        * 100
        if eligible_internal_count
        else 0
    )

    amount_match_rate = (
        matched_amount
        / valid_internal["NOMINAL_MATCH"].sum()
        * 100
        if valid_internal["NOMINAL_MATCH"].sum()
        else 0
    )

    expected_amount = (
        valid_internal[
            "NOMINAL_MATCH"
        ].sum()
    )

    actual_amount = (
        valid_bank[
            "NOMINAL_MATCH"
        ].sum()
    )

    amount_difference = (
        actual_amount
        - expected_amount
    )

    summary = {
        "reconciliation_id": reconciliation_id,

        "total_internal_success":
            total_internal_success,

        "total_internal_success_amount":
            total_internal_success_amount,

        "internal_invalid_va_count":
            len(invalid_internal),

        "internal_invalid_va_amount":
            invalid_internal[
                "NOMINAL_ORIGINAL"
            ].sum()
            if not invalid_internal.empty
            else 0,

        "total_bank_credit":
            total_bank_credit,

        "total_bank_credit_amount":
            total_bank_credit_amount,

        "bank_invalid_va_count":
            len(invalid_bank),

        "bank_invalid_va_amount":
            invalid_bank[
                "NOMINAL_KREDIT_NUM"
            ].sum()
            if not invalid_bank.empty
            else 0,

        "eligible_internal_count":
            eligible_internal_count,

        "eligible_bank_count":
            eligible_bank_count,

        "matched_count":
            matched_count,

        "matched_amount":
            matched_amount,

        "issue_fmss_count":
            len(only_int),

        "issue_fmss_amount":
            issue_fmss_amount,

        "issue_bank_count":
            len(only_bnk),

        "issue_bank_amount":
            issue_bank_amount,

        "expected_amount":
            expected_amount,

        "actual_amount":
            actual_amount,

        "amount_difference":
            amount_difference,

        "transaction_match_rate":
            transaction_match_rate,

        "amount_match_rate":
            amount_match_rate,

        "carry_over_count":
            carry_count,
    }

    return (
        matched,
        only_int,
        only_bnk,
        invalid_internal,
        invalid_bank,
        summary
    )


# ============================================================
# SESSION STATE
# ============================================================

for key, default in [

    ("sudah_diproses", False),

    ("df_matched", pd.DataFrame()),

    ("df_selisih_int", pd.DataFrame()),

    ("df_selisih_bnk", pd.DataFrame()),

    ("df_invalid_int", pd.DataFrame()),

    ("df_invalid_bnk", pd.DataFrame()),

    ("summary", {}),

    ("pilihan_bank_terakhir", ""),

]:
    if key not in st.session_state:
        st.session_state[key] = default


# ============================================================
# 1. PENGATURAN
# ============================================================

st.subheader("1. Pengaturan Data")

opsi_bank = [
    ""
] + list(
    BANK_CONFIGS.keys()
)

pilihan_bank = st.selectbox(
    "Pilih Bank Sumber Mutasi:",
    opsi_bank
)

if (
    pilihan_bank
    != st.session_state[
        "pilihan_bank_terakhir"
    ]
):

    st.session_state[
        "sudah_diproses"
    ] = False

    st.session_state[
        "pilihan_bank_terakhir"
    ] = pilihan_bank


mode_dev = st.checkbox(
    "Tampilkan detail error teknis "
    "(mode developer)",
    value=False
)


fee_override = None
fee_reason = ""


if (
    pilihan_bank
    and BANK_CONFIGS
    .get(pilihan_bank, {})
    .get("configured")
):

    config_preview = BANK_CONFIGS[
        pilihan_bank
    ]

    default_fee = int(
        config_preview[
            "fee_adjustment"
        ]
    )

    fee_override = st.number_input(
        "Fee/selisih admin "
        f"{pilihan_bank}",
        value=default_fee,
        step=500,
        min_value=0,
        help=(
            "Nominal internal + fee "
            "= nominal bank."
        )
    )

    if fee_override != default_fee:

        fee_reason = st.text_input(
            "Alasan perubahan fee",
            placeholder=(
                "Contoh: Tarif admin periode "
                "Agustus 2026"
            )
        )

        if not fee_reason.strip():

            st.warning(
                "⚠️ Fee berbeda dari default. "
                "Isi alasan perubahan sebelum "
                "melakukan rekonsiliasi."
            )


# ============================================================
# 2. UPLOAD
# ============================================================

st.subheader("2. Unggah File")

col1, col2 = st.columns(2)

with col1:

    file_int = st.file_uploader(
        "Unggah CSV/XLSX Dari FMSS",
        type=["csv", "xlsx"],
        key="int"
    )

with col2:

    label_bank = (
        f"Unggah CSV/XLSX "
        f"Mutasi Bank ({pilihan_bank})"
        if pilihan_bank
        else
        "Unggah CSV/XLSX Mutasi Bank"
    )

    file_bnk = st.file_uploader(
        label_bank,
        type=["csv", "xlsx"],
        key="bnk"
    )


with st.expander(
    "📤 Selisih FMSS dari proses sebelumnya "
    "(opsional - cutoff H+1)"
):

    st.caption(
        "Gunakan ISSUE_FMSS dari proses sebelumnya "
        "jika bank menggunakan settlement/cutoff H+1."
    )

    file_carry = st.file_uploader(
        "Unggah ISSUE_FMSS sebelumnya",
        type=["csv", "xlsx"],
        key="carry"
    )


# ============================================================
# 3. PROCESS
# ============================================================

if (
    pilihan_bank
    and file_int
    and file_bnk
):

    st.divider()

    config = BANK_CONFIGS[
        pilihan_bank
    ]

    if not config.get("configured"):

        st.warning(
            f"🚧 Modul {pilihan_bank} "
            "belum dikonfigurasi."
        )

    else:

        if st.button(
            f"🚀 Mulai Croscek Data "
            f"{pilihan_bank}",
            type="primary"
        ):

            if (
                fee_override
                != config["fee_adjustment"]
                and not fee_reason.strip()
            ):

                st.error(
                    "❌ Fee berbeda dari default "
                    "tetapi alasan perubahan belum diisi."
                )

            else:

                try:

                    with st.spinner(
                        "Sedang memproses "
                        "rekonsiliasi..."
                    ):

                        recon_id = (
                            generate_reconciliation_id()
                        )

                        df_int_raw = read_any(
                            file_int
                        )

                        df_bnk_raw = read_any(
                            file_bnk
                        )

                        df_carry_raw = (
                            read_any(file_carry)
                            if file_carry
                            else None
                        )

                        active_config = dict(
                            config
                        )

                        if fee_override is not None:

                            active_config[
                                "fee_adjustment"
                            ] = fee_override

                        (
                            matched,
                            only_int,
                            only_bnk,
                            invalid_int,
                            invalid_bnk,
                            summary
                        ) = process_reconciliation(

                            df_int_raw,
                            df_bnk_raw,
                            active_config,
                            recon_id,
                            df_carry_raw
                        )

                        summary[
                            "bank"
                        ] = pilihan_bank

                        summary[
                            "processing_time"
                        ] = datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )

                        summary[
                            "fee_adjustment"
                        ] = fee_override

                        summary[
                            "fee_reason"
                        ] = (
                            fee_reason
                            if fee_override
                            != config[
                                "fee_adjustment"
                            ]
                            else
                            "DEFAULT_CONFIG"
                        )

                        st.session_state[
                            "df_matched"
                        ] = matched

                        st.session_state[
                            "df_selisih_int"
                        ] = only_int

                        st.session_state[
                            "df_selisih_bnk"
                        ] = only_bnk

                        st.session_state[
                            "df_invalid_int"
                        ] = invalid_int

                        st.session_state[
                            "df_invalid_bnk"
                        ] = invalid_bnk

                        st.session_state[
                            "summary"
                        ] = summary

                        st.session_state[
                            "sudah_diproses"
                        ] = True

                except ValueError as e:

                    st.error(
                        f"❌ {e}"
                    )

                    st.session_state[
                        "sudah_diproses"
                    ] = False

                except Exception as e:

                    st.error(
                        "❌ Terjadi kesalahan "
                        "tak terduga saat "
                        "memproses data."
                    )

                    if mode_dev:

                        st.code(
                            traceback.format_exc()
                        )

                    else:

                        st.caption(
                            f"Technical detail: {e}"
                        )

                    st.session_state[
                        "sudah_diproses"
                    ] = False


# ============================================================
# 4. RESULT
# ============================================================

if st.session_state[
    "sudah_diproses"
]:

    df_matched = st.session_state[
        "df_matched"
    ]

    df_selisih_int = st.session_state[
        "df_selisih_int"
    ]

    df_selisih_bnk = st.session_state[
        "df_selisih_bnk"
    ]

    df_invalid_int = st.session_state[
        "df_invalid_int"
    ]

    df_invalid_bnk = st.session_state[
        "df_invalid_bnk"
    ]

    summary = st.session_state[
        "summary"
    ]

    st.subheader(
        f"🎯 Ringkasan Rekonsiliasi "
        f"{pilihan_bank}"
    )

    # --------------------------------------------------------
    # KPI TRANSACTION
    # --------------------------------------------------------

    k1, k2, k3, k4 = st.columns(4)

    k1.metric(
        "✅ Matched",
        f"{summary['matched_count']:,} Trx"
    )

    k2.metric(
        "⚠️ FMSS Issue",
        f"{summary['issue_fmss_count']:,} Trx"
    )

    k3.metric(
        "⚠️ Bank Issue",
        f"{summary['issue_bank_count']:,} Trx"
    )

    k4.metric(
        "🎯 Match Rate",
        f"{summary['transaction_match_rate']:.2f}%"
    )


    # --------------------------------------------------------
    # FINANCIAL KPI
    # --------------------------------------------------------

    st.markdown(
        "### 💰 Financial Reconciliation"
    )

    f1, f2, f3, f4 = st.columns(4)

    f1.metric(
        "Expected Bank Amount",
        f"Rp {summary['expected_amount']:,.0f}"
    )

    f2.metric(
        "Actual Bank Amount",
        f"Rp {summary['actual_amount']:,.0f}"
    )

    f3.metric(
        "Matched Amount",
        f"Rp {summary['matched_amount']:,.0f}"
    )

    f4.metric(
        "Amount Difference",
        f"Rp {summary['amount_difference']:,.0f}"
    )


    # --------------------------------------------------------
    # MATCH RATE
    # --------------------------------------------------------

    r1, r2 = st.columns(2)

    r1.metric(
        "📊 Transaction Match Rate",
        f"{summary['transaction_match_rate']:.2f}%"
    )

    r2.metric(
        "💰 Amount Match Rate",
        f"{summary['amount_match_rate']:.2f}%"
    )


    # --------------------------------------------------------
    # HEALTH STATUS
    # --------------------------------------------------------

    match_rate = (
        summary["transaction_match_rate"]
    )

    amount_rate = (
        summary["amount_match_rate"]
    )

    if (
        match_rate >= 99.5
        and amount_rate >= 99.5
        and summary["internal_invalid_va_count"] == 0
        and summary["bank_invalid_va_count"] == 0
    ):

        st.success(
            "🟢 RECONCILIATION HEALTHY — "
            "Match rate dan amount reconciliation "
            "berada dalam kondisi sangat baik."
        )

    elif (
        match_rate >= 98
        and amount_rate >= 98
    ):

        st.warning(
            "🟡 RECONCILIATION WARNING — "
            "Masih terdapat exception yang "
            "perlu ditindaklanjuti."
        )

    else:

        st.error(
            "🔴 RECONCILIATION CRITICAL — "
            "Terdapat selisih signifikan yang "
            "perlu segera diinvestigasi."
        )


    # --------------------------------------------------------
    # DATA QUALITY
    # --------------------------------------------------------

    with st.expander(
        "🔎 Data Quality & Completeness"
    ):

        q1, q2, q3, q4 = st.columns(4)

        q1.metric(
            "FMSS SUCCESS",
            f"{summary['total_internal_success']:,}"
        )

        q2.metric(
            "FMSS Invalid VA",
            f"{summary['internal_invalid_va_count']:,}"
        )

        q3.metric(
            "Bank Credit",
            f"{summary['total_bank_credit']:,}"
        )

        q4.metric(
            "Bank Invalid VA",
            f"{summary['bank_invalid_va_count']:,}"
        )


        st.write(
            "### Nominal Invalid VA"
        )

        iq1, iq2 = st.columns(2)

        iq1.metric(
            "FMSS Invalid VA Amount",
            f"Rp {summary['internal_invalid_va_amount']:,.0f}"
        )

        iq2.metric(
            "Bank Invalid VA Amount",
            f"Rp {summary['bank_invalid_va_amount']:,.0f}"
        )


    # --------------------------------------------------------
    # CUTOFF
    # --------------------------------------------------------

    if (
        "MATCH_TIMING"
        in df_matched.columns
    ):

        timing_counts = (
            df_matched[
                "MATCH_TIMING"
            ]
            .value_counts()
        )

        st.markdown(
            "### 🕒 Settlement / Cutoff"
        )

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Same Day",
            f"{timing_counts.get('SAME_DAY', 0):,}"
        )

        c2.metric(
            "H+1",
            f"{timing_counts.get('H+1', 0):,}"
        )

        c3.metric(
            "H+2+",
            f"{timing_counts.get('H+2+', 0):,}"
        )


    st.divider()


    # ========================================================
    # ISSUE FMSS
    # ========================================================

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            "#### 🚨 Issue FMSS"
        )

        if not df_selisih_int.empty:

            cols = [
                "VA_CODE",
                "INT__nominal",
                "ISSUE_TYPE"
            ]

            if (
                "INT__SUMBER_DATA"
                in df_selisih_int.columns
            ):

                cols.append(
                    "INT__SUMBER_DATA"
                )

            display_int = (
                df_selisih_int[
                    cols
                ].rename(
                    columns={
                        "VA_CODE": "KODE VA",
                        "INT__nominal": "NOMINAL",
                        "ISSUE_TYPE": "ISSUE",
                        "INT__SUMBER_DATA": "SUMBER"
                    }
                )
            )

            st.dataframe(
                display_int,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue FMSS."
            )


    # ========================================================
    # ISSUE BANK
    # ========================================================

    with col2:

        st.markdown(
            "#### 🚨 Issue Bank"
        )

        if not df_selisih_bnk.empty:

            bank_credit_col_prefixed = (
                f"BNK__{config['bank_credit_col']}"
            )

            display_bank = (
                df_selisih_bnk[
                    [
                        "VA_CODE",
                        bank_credit_col_prefixed,
                        "ISSUE_TYPE"
                    ]
                ].rename(
                    columns={
                        "VA_CODE": "KODE VA",
                        bank_credit_col_prefixed:
                            "NOMINAL",
                        "ISSUE_TYPE": "ISSUE"
                    }
                )
            )

            st.dataframe(
                display_bank,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue Bank."
            )


    # ========================================================
    # INVALID VA
    # ========================================================

    with st.expander(
        "⚠️ Transaksi dengan VA Tidak Teridentifikasi"
    ):

        iv1, iv2 = st.columns(2)

        with iv1:

            st.markdown(
                "##### FMSS Invalid VA"
            )

            if not df_invalid_int.empty:

                st.dataframe(
                    df_invalid_int,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.success(
                    "Tidak ada FMSS invalid VA."
                )


        with iv2:

            st.markdown(
                "##### Bank Invalid VA"
            )

            if not df_invalid_bnk.empty:

                st.dataframe(
                    df_invalid_bnk,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.success(
                    "Tidak ada Bank invalid VA."
                )


    # ========================================================
    # EXPORT
    # ========================================================

    st.divider()

    st.markdown(
        "### 📥 Download Laporan"
    )

    drop_cols = [
        "NOMINAL_MATCH",
        "_occ",
        "_merge"
    ]

    issue_fmss_export = (
        df_selisih_int
        .drop(
            columns=drop_cols,
            errors="ignore"
        )
        .copy()
    )

    issue_fmss_export.columns = [
        c.replace(
            "INT__",
            ""
        )
        for c in issue_fmss_export.columns
    ]


    matched_export = (
        df_matched
        .drop(
            columns=drop_cols,
            errors="ignore"
        )
        .copy()
    )


    issue_bank_export = (
        df_selisih_bnk
        .drop(
            columns=drop_cols,
            errors="ignore"
        )
        .copy()
    )


    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ):

        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        summary_df = pd.DataFrame([
            {
                "RECONCILIATION_ID":
                    summary["reconciliation_id"],

                "BANK":
                    summary["bank"],

                "PROCESSING_TIME":
                    summary["processing_time"],

                "EXPECTED_AMOUNT":
                    summary["expected_amount"],

                "ACTUAL_BANK_AMOUNT":
                    summary["actual_amount"],

                "MATCHED_AMOUNT":
                    summary["matched_amount"],

                "AMOUNT_DIFFERENCE":
                    summary["amount_difference"],

                "TRANSACTION_MATCH_RATE":
                    summary["transaction_match_rate"],

                "AMOUNT_MATCH_RATE":
                    summary["amount_match_rate"],

                "FMSS_ISSUE_COUNT":
                    summary["issue_fmss_count"],

                "FMSS_ISSUE_AMOUNT":
                    summary["issue_fmss_amount"],

                "BANK_ISSUE_COUNT":
                    summary["issue_bank_count"],

                "BANK_ISSUE_AMOUNT":
                    summary["issue_bank_amount"],

                "FMSS_INVALID_VA":
                    summary["internal_invalid_va_count"],

                "BANK_INVALID_VA":
                    summary["bank_invalid_va_count"],

                "CARRY_OVER_COUNT":
                    summary["carry_over_count"],

                "FEE_ADJUSTMENT":
                    summary["fee_adjustment"],

                "FEE_REASON":
                    summary["fee_reason"],
            }
        ])

        summary_df.to_excel(
            writer,
            sheet_name="SUMMARY",
            index=False
        )


        # ----------------------------------------------------
        # ISSUE FMSS
        # ----------------------------------------------------

        if not issue_fmss_export.empty:

            issue_fmss_export.to_excel(
                writer,
                sheet_name="ISSUE_FMSS",
                index=False
            )

        else:

            pd.DataFrame({
                "Info": [
                    "Bersih! Tidak ada selisih FMSS."
                ]
            }).to_excel(
                writer,
                sheet_name="ISSUE_FMSS",
                index=False
            )


        # ----------------------------------------------------
        # ISSUE BANK
        # ----------------------------------------------------

        if not issue_bank_export.empty:

            issue_bank_export.to_excel(
                writer,
                sheet_name="ISSUE_BANK",
                index=False
            )

        else:

            pd.DataFrame({
                "Info": [
                    "Bersih! Tidak ada selisih Bank."
                ]
            }).to_excel(
                writer,
                sheet_name="ISSUE_BANK",
                index=False
            )


        # ----------------------------------------------------
        # MATCHED
        # ----------------------------------------------------

        if not matched_export.empty:

            matched_export.to_excel(
                writer,
                sheet_name="MATCHED_OK",
                index=False
            )

        else:

            pd.DataFrame({
                "Info": [
                    "Tidak ada data matched."
                ]
            }).to_excel(
                writer,
                sheet_name="MATCHED_OK",
                index=False
            )


        # ----------------------------------------------------
        # INVALID VA
        # ----------------------------------------------------

        if not df_invalid_int.empty:

            df_invalid_int.to_excel(
                writer,
                sheet_name="INVALID_VA_FMSS",
                index=False
            )

        if not df_invalid_bnk.empty:

            df_invalid_bnk.to_excel(
                writer,
                sheet_name="INVALID_VA_BANK",
                index=False
            )


    st.download_button(

        label=(
            "📥 Download Laporan "
            "Rekonsiliasi Lengkap (.xlsx)"
        ),

        data=output.getvalue(),

        file_name=(
            f"Laporan_Rekonsiliasi_"
            f"{pilihan_bank}_"
            f"{summary['reconciliation_id']}.xlsx"
        ),

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        type="primary"
    )


    st.caption(
        "💡 Sheet ISSUE_FMSS dapat digunakan "
        "sebagai carry-over pada proses "
        "rekonsiliasi berikutnya."
    )


elif (
    pilihan_bank == ""
    and (file_int or file_bnk)
):

    st.info(
        "💡 Silakan pilih Bank Sumber Mutasi "
        "terlebih dahulu."
    )
