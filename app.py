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
# BANK CONFIGURATION
# ============================================================

BANK_CONFIGS = {

    "BRIVA": {

        "configured": True,

        # ----------------------------------------------------
        # FMSS:
        # 57888 = BRIVA Fastpay
        # 57708 = BRIVA Rajabiller
        # ----------------------------------------------------
        "internal_extractor": {
            "fn": "multi_prefix",
            "source_col": "keterangan",
            "prefixes": {
                "57888": "BRIVA FASTPAY",
                "57708": "BRIVA RAJABILLER"
            }
        },

        # ----------------------------------------------------
        # BANK:
        # Untuk mutasi BRIVA Fastpay gunakan 57888
        # ----------------------------------------------------
        "bank_extractor": {
            "fn": "multi_prefix",
            "source_col": "DESK_TRAN",
            "prefixes": {
                "57888": "BRIVA FASTPAY"
            }
        },

        "bank_credit_col": "MUTASI_KREDIT",

        # Nominal FMSS + Rp1.000 = nominal bank
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
# SESSION STATE
# ============================================================

DEFAULT_SESSION = {

    "sudah_diproses": False,

    "df_matched":
        pd.DataFrame(),

    "df_selisih_int":
        pd.DataFrame(),

    "df_selisih_bnk":
        pd.DataFrame(),

    "df_invalid_int":
        pd.DataFrame(),

    "df_invalid_bnk":
        pd.DataFrame(),

    "summary":
        {},

    "pilihan_bank_terakhir":
        "",
}


for key, default in DEFAULT_SESSION.items():

    if key not in st.session_state:

        st.session_state[key] = default


# ============================================================
# VA EXTRACTOR
# ============================================================

def extract_va_multi_prefix(
    text,
    prefixes
):

    if pd.isna(text):
        return None, None

    text = str(text)

    for prefix, va_type in prefixes.items():

        match = re.search(
            rf'({re.escape(prefix)}\d{{5,15}})',
            text
        )

        if match:

            return (
                match.group(1),
                va_type
            )

    return (
        None,
        None
    )


def apply_va_extractor(
    df,
    config
):

    source_col = config[
        "source_col"
    ]

    prefixes = config[
        "prefixes"
    ]

    extracted = df[
        source_col
    ].apply(
        lambda x:
            extract_va_multi_prefix(
                x,
                prefixes
            )
    )

    df = df.copy()

    df[
        "VA_CODE"
    ] = extracted.apply(
        lambda x: x[0]
    )

    df[
        "VA_TYPE"
    ] = extracted.apply(
        lambda x: x[1]
    )

    return df


# ============================================================
# NOMINAL PARSER
# ============================================================

def parse_nominal(
    series
):

    if pd.api.types.is_numeric_dtype(
        series
    ):

        return series.fillna(0)

    cleaned = (
        series
        .astype(str)
        .str.replace(
            r"Rp\.?",
            "",
            regex=True
        )
        .str.replace(
            " ",
            "",
            regex=False
        )
        .str.replace(
            ".",
            "",
            regex=False
        )
        .str.replace(
            ",",
            ".",
            regex=False
        )
        .str.strip()
    )

    return pd.to_numeric(
        cleaned,
        errors="coerce"
    ).fillna(0)


# ============================================================
# READ FILE
# ============================================================

def read_any(
    uploaded_file
):

    uploaded_file.seek(0)

    filename = (
        uploaded_file.name
        .lower()
    )

    if filename.endswith(
        ".csv"
    ):

        return pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python"
        )

    return pd.read_excel(
        uploaded_file
    )


# ============================================================
# VALIDATE COLUMNS
# ============================================================

def validate_columns(
    df,
    required_cols,
    source_label
):

    missing = [
        col
        for col in required_cols
        if col not in df.columns
    ]

    if missing:

        raise ValueError(
            f"File {source_label} tidak memiliki "
            f"kolom wajib: {', '.join(missing)}. "
            f"Kolom tersedia: "
            f"{', '.join(map(str, df.columns))}"
        )


# ============================================================
# RECONCILIATION ID
# ============================================================

def generate_reconciliation_id():

    return datetime.now().strftime(
        "REC-%Y%m%d-%H%M%S"
    )


# ============================================================
# 1 TO 1 MATCHING ENGINE
# ============================================================

def tally_1_to_1(
    df_int,
    df_bnk
):

    di = df_int.copy()
    db = df_bnk.copy()

    # --------------------------------------------------------
    # Duplicate counter
    # --------------------------------------------------------

    di["_occ"] = (
        di
        .groupby(
            [
                "VA_CODE",
                "NOMINAL_MATCH"
            ]
        )
        .cumcount()
    )

    db["_occ"] = (
        db
        .groupby(
            [
                "VA_CODE",
                "NOMINAL_MATCH"
            ]
        )
        .cumcount()
    )

    # --------------------------------------------------------
    # Prefix columns
    # --------------------------------------------------------

    di = di.add_prefix(
        "INT__"
    )

    db = db.add_prefix(
        "BNK__"
    )

    # --------------------------------------------------------
    # Restore matching keys
    # --------------------------------------------------------

    di = di.rename(
        columns={
            "INT__VA_CODE":
                "VA_CODE",

            "INT__NOMINAL_MATCH":
                "NOMINAL_MATCH",

            "INT___occ":
                "_occ"
        }
    )

    db = db.rename(
        columns={
            "BNK__VA_CODE":
                "VA_CODE",

            "BNK__NOMINAL_MATCH":
                "NOMINAL_MATCH",

            "BNK___occ":
                "_occ"
        }
    )

    # --------------------------------------------------------
    # Outer merge
    # --------------------------------------------------------

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

    return (
        matched,
        only_int,
        only_bnk
    )


# ============================================================
# CLASSIFY UNMATCHED
# ============================================================

def classify_unmatched(
    only_int,
    only_bnk
):

    only_int = only_int.copy()
    only_bnk = only_bnk.copy()

    bank_remaining_va = set(
        only_bnk[
            "VA_CODE"
        ]
        .dropna()
    )

    int_remaining_va = set(
        only_int[
            "VA_CODE"
        ]
        .dropna()
    )

    # --------------------------------------------------------
    # FMSS
    # --------------------------------------------------------

    if not only_int.empty:

        only_int[
            "ISSUE_TYPE"
        ] = (
            only_int[
                "VA_CODE"
            ]
            .isin(
                bank_remaining_va
            )
            .map(
                {
                    True:
                        "AMOUNT_MISMATCH",

                    False:
                        "FMSS_ONLY"
                }
            )
        )

    # --------------------------------------------------------
    # BANK
    # --------------------------------------------------------

    if not only_bnk.empty:

        only_bnk[
            "ISSUE_TYPE"
        ] = (
            only_bnk[
                "VA_CODE"
            ]
            .isin(
                int_remaining_va
            )
            .map(
                {
                    True:
                        "AMOUNT_MISMATCH",

                    False:
                        "BANK_ONLY"
                }
            )
        )

    return (
        only_int,
        only_bnk
    )


# ============================================================
# MAIN RECONCILIATION ENGINE
# ============================================================

def process_reconciliation(
    df_int_raw,
    df_bnk_raw,
    config,
    reconciliation_id,
    df_carry_over=None
):

    # ========================================================
    # VALIDATE INTERNAL
    # ========================================================

    validate_columns(
        df_int_raw,
        INTERNAL_REQUIRED_COLS,
        "FMSS (Internal)"
    )

    # ========================================================
    # VALIDATE BANK
    # ========================================================

    bank_required = [
        config[
            "bank_extractor"
        ][
            "source_col"
        ],

        config[
            "bank_credit_col"
        ]
    ]

    validate_columns(
        df_bnk_raw,
        bank_required,
        "Mutasi Bank"
    )

    # ========================================================
    # INTERNAL SUCCESS
    # ========================================================

    df_int_success = (
        df_int_raw[
            df_int_raw[
                "status"
            ]
            .astype(str)
            .str.upper()
            == "SUKSES"
        ]
        .copy()
    )

    total_internal_success = (
        len(
            df_int_success
        )
    )

    df_int_success[
        "SUMBER_DATA"
    ] = "BARU"

    # ========================================================
    # CARRY OVER
    # ========================================================

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

        df_carry = (
            df_carry_over.copy()
        )

        df_carry[
            "SUMBER_DATA"
        ] = "CARRY_OVER"

        carry_count = len(
            df_carry
        )

        df_int_success = pd.concat(
            [
                df_carry,
                df_int_success
            ],
            ignore_index=True
        )

    # ========================================================
    # EXTRACT FMSS VA
    # ========================================================

    df_int_success = (
        apply_va_extractor(
            df_int_success,
            config[
                "internal_extractor"
            ]
        )
    )

    # ========================================================
    # FMSS NOMINAL
    # ========================================================

    df_int_success[
        "NOMINAL_ORIGINAL"
    ] = parse_nominal(
        df_int_success[
            "nominal"
        ]
    )

    df_int_success[
        "NOMINAL_MATCH"
    ] = (
        df_int_success[
            "NOMINAL_ORIGINAL"
        ]
        +
        config[
            "fee_adjustment"
        ]
    )

    # ========================================================
    # IMPORTANT
    #
    # INVALID VA hanya benar-benar tidak memiliki
    # prefix 57888 ATAU 57708.
    # ========================================================

    invalid_internal = (
        df_int_success[
            df_int_success[
                "VA_CODE"
            ].isna()
        ]
        .copy()
    )

    # ========================================================
    # VALID FMSS
    #
    # 57888 = Fastpay
    # 57708 = Rajabiller
    # ========================================================

    valid_internal = (
        df_int_success[
            df_int_success[
                "VA_CODE"
            ].notna()
        ]
        .copy()
        .reset_index(
            drop=True
        )
    )

    # ========================================================
    # BANK CREDIT
    # ========================================================

    df_bnk = (
        df_bnk_raw.copy()
    )

    df_bnk[
        "NOMINAL_KREDIT_NUM"
    ] = parse_nominal(
        df_bnk[
            config[
                "bank_credit_col"
            ]
        ]
    )

    df_bnk_credit = (
        df_bnk[
            df_bnk[
                "NOMINAL_KREDIT_NUM"
            ] > 0
        ]
        .copy()
    )

    total_bank_credit = (
        len(
            df_bnk_credit
        )
    )

    # ========================================================
    # EXTRACT BANK VA
    # ========================================================

    df_bnk_credit = (
        apply_va_extractor(
            df_bnk_credit,
            config[
                "bank_extractor"
            ]
        )
    )

    # ========================================================
    # BANK INVALID VA
    # ========================================================

    invalid_bank = (
        df_bnk_credit[
            df_bnk_credit[
                "VA_CODE"
            ].isna()
        ]
        .copy()
    )

    # ========================================================
    # VALID BANK
    # ========================================================

    valid_bank = (
        df_bnk_credit[
            df_bnk_credit[
                "VA_CODE"
            ].notna()
        ]
        .copy()
        .reset_index(
            drop=True
        )
    )

    valid_bank[
        "NOMINAL_MATCH"
    ] = (
        valid_bank[
            "NOMINAL_KREDIT_NUM"
        ]
    )

    # ========================================================
    # DATA QUALITY
    # ========================================================

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

    # ========================================================
    # 1 TO 1 MATCH
    # ========================================================

    (
        matched,
        only_int,
        only_bnk
    ) = tally_1_to_1(
        valid_internal,
        valid_bank
    )

    # ========================================================
    # CLASSIFICATION
    # ========================================================

    (
        only_int,
        only_bnk
    ) = classify_unmatched(
        only_int,
        only_bnk
    )

    # ========================================================
    # MATCH METADATA
    # ========================================================

    if not matched.empty:

        matched[
            "MATCH_STATUS"
        ] = "MATCHED"

        matched[
            "MATCH_AMOUNT"
        ] = (
            matched[
                "NOMINAL_MATCH"
            ]
        )

        matched[
            "RECONCILIATION_ID"
        ] = reconciliation_id

    # ========================================================
    # ISSUE METADATA
    # ========================================================

    if not only_int.empty:

        only_int[
            "RECONCILIATION_ID"
        ] = reconciliation_id

    if not only_bnk.empty:

        only_bnk[
            "RECONCILIATION_ID"
        ] = reconciliation_id

    # ========================================================
    # SUMMARY
    # ========================================================

    matched_count = (
        len(
            matched
        )
    )

    matched_amount = (
        matched[
            "NOMINAL_MATCH"
        ].sum()
        if not matched.empty
        else 0
    )

    issue_fmss_amount = (
        only_int[
            "NOMINAL_MATCH"
        ].sum()
        if not only_int.empty
        else 0
    )

    issue_bank_amount = (
        only_bnk[
            "NOMINAL_MATCH"
        ].sum()
        if not only_bnk.empty
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

    eligible_internal_count = (
        len(
            valid_internal
        )
    )

    transaction_match_rate = (
        matched_count
        /
        eligible_internal_count
        *
        100
        if eligible_internal_count
        else 0
    )

    amount_match_rate = (
        matched_amount
        /
        expected_amount
        *
        100
        if expected_amount
        else 0
    )

    summary = {

        "reconciliation_id":
            reconciliation_id,

        "bank":
            "BRIVA",

        "processing_time":
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

        "total_internal_success":
            total_internal_success,

        "internal_invalid_va_count":
            len(
                invalid_internal
            ),

        "internal_invalid_va_amount":
            (
                invalid_internal[
                    "NOMINAL_ORIGINAL"
                ].sum()
                if not invalid_internal.empty
                else 0
            ),

        "total_bank_credit":
            total_bank_credit,

        "bank_invalid_va_count":
            len(
                invalid_bank
            ),

        "bank_invalid_va_amount":
            (
                invalid_bank[
                    "NOMINAL_KREDIT_NUM"
                ].sum()
                if not invalid_bank.empty
                else 0
            ),

        "eligible_internal_count":
            eligible_internal_count,

        "eligible_bank_count":
            len(
                valid_bank
            ),

        "matched_count":
            matched_count,

        "matched_amount":
            matched_amount,

        "issue_fmss_count":
            len(
                only_int
            ),

        "issue_fmss_amount":
            issue_fmss_amount,

        "issue_bank_count":
            len(
                only_bnk
            ),

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

        "fee_adjustment":
            config[
                "fee_adjustment"
            ]
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
# UI - PENGATURAN
# ============================================================

st.subheader(
    "1. Pengaturan Data"
)

opsi_bank = [
    ""
] + list(
    BANK_CONFIGS.keys()
)

pilihan_bank = st.selectbox(
    "Pilih Bank Sumber Mutasi:",
    opsi_bank
)


# ============================================================
# RESET STATE
# ============================================================

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


# ============================================================
# DEVELOPER MODE
# ============================================================

mode_dev = st.checkbox(
    "Tampilkan detail error teknis",
    value=False
)


# ============================================================
# FEE
# ============================================================

fee_override = None
fee_reason = ""

if pilihan_bank:

    config_preview = (
        BANK_CONFIGS[
            pilihan_bank
        ]
    )

    if config_preview.get(
        "configured"
    ):

        default_fee = int(
            config_preview[
                "fee_adjustment"
            ]
        )

        fee_override = st.number_input(
            "Fee / selisih admin",
            value=default_fee,
            step=500,
            min_value=0
        )

        if (
            fee_override
            != default_fee
        ):

            fee_reason = st.text_input(
                "Alasan perubahan fee"
            )


# ============================================================
# UPLOAD
# ============================================================

st.subheader(
    "2. Unggah File"
)

col1, col2 = st.columns(2)

with col1:

    file_int = st.file_uploader(
        "Unggah CSV/XLSX Dari FMSS",
        type=[
            "csv",
            "xlsx"
        ],
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
        type=[
            "csv",
            "xlsx"
        ],
        key="bnk"
    )


# ============================================================
# CARRY OVER
# ============================================================

with st.expander(
    "📤 Carry-over ISSUE_FMSS sebelumnya"
):

    st.caption(
        "Opsional. Digunakan apabila terdapat "
        "transaksi FMSS yang baru menerima "
        "mutasi bank pada proses berikutnya."
    )

    file_carry = st.file_uploader(
        "Upload ISSUE_FMSS sebelumnya",
        type=[
            "csv",
            "xlsx"
        ],
        key="carry"
    )


# ============================================================
# PROCESS
# ============================================================

if (
    pilihan_bank
    and file_int
    and file_bnk
):

    st.divider()

    config = (
        BANK_CONFIGS[
            pilihan_bank
        ]
    )

    if not config.get(
        "configured"
    ):

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
                != config[
                    "fee_adjustment"
                ]
                and not fee_reason.strip()
            ):

                st.error(
                    "Fee berbeda dari default. "
                    "Isi alasan perubahan terlebih dahulu."
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

                        df_int_raw = (
                            read_any(
                                file_int
                            )
                        )

                        df_bnk_raw = (
                            read_any(
                                file_bnk
                            )
                        )

                        df_carry_raw = (
                            read_any(
                                file_carry
                            )
                            if file_carry
                            else None
                        )

                        active_config = (
                            config.copy()
                        )

                        if (
                            fee_override
                            is not None
                        ):

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
                            "fee_reason"
                        ] = (
                            fee_reason
                            if fee_reason
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

                except Exception as e:

                    st.error(
                        "❌ Terjadi kesalahan "
                        "saat memproses data."
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
# RESULT
# ============================================================

if st.session_state[
    "sudah_diproses"
]:

    df_matched = (
        st.session_state[
            "df_matched"
        ]
    )

    df_selisih_int = (
        st.session_state[
            "df_selisih_int"
        ]
    )

    df_selisih_bnk = (
        st.session_state[
            "df_selisih_bnk"
        ]
    )

    df_invalid_int = (
        st.session_state[
            "df_invalid_int"
        ]
    )

    df_invalid_bnk = (
        st.session_state[
            "df_invalid_bnk"
        ]
    )

    summary = (
        st.session_state[
            "summary"
        ]
    )


    # ========================================================
    # SUMMARY
    # ========================================================

    st.subheader(
        f"🎯 Ringkasan Rekonsiliasi "
        f"{pilihan_bank}"
    )

    k1, k2, k3, k4 = st.columns(4)

    k1.metric(
        "✅ Matched",
        f"{summary['matched_count']:,} Trx"
    )

    k2.metric(
        "⚠️ Issue FMSS",
        f"{summary['issue_fmss_count']:,} Trx"
    )

    k3.metric(
        "⚠️ Issue Bank",
        f"{summary['issue_bank_count']:,} Trx"
    )

    k4.metric(
        "🎯 Match Rate",
        f"{summary['transaction_match_rate']:.2f}%"
    )


    # ========================================================
    # FINANCIAL
    # ========================================================

    st.markdown(
        "### 💰 Financial Reconciliation"
    )

    f1, f2, f3, f4 = st.columns(4)

    f1.metric(
        "Expected",
        f"Rp {summary['expected_amount']:,.0f}"
    )

    f2.metric(
        "Actual Bank",
        f"Rp {summary['actual_amount']:,.0f}"
    )

    f3.metric(
        "Matched",
        f"Rp {summary['matched_amount']:,.0f}"
    )

    f4.metric(
        "Difference",
        f"Rp {summary['amount_difference']:,.0f}"
    )


    # ========================================================
    # AMOUNT RATE
    # ========================================================

    r1, r2 = st.columns(2)

    r1.metric(
        "📊 Transaction Match Rate",
        f"{summary['transaction_match_rate']:.2f}%"
    )

    r2.metric(
        "💰 Amount Match Rate",
        f"{summary['amount_match_rate']:.2f}%"
    )


    # ========================================================
    # HEALTH
    # ========================================================

    if (
        summary[
            "transaction_match_rate"
        ] >= 99.5
        and
        summary[
            "amount_match_rate"
        ] >= 99.5
        and
        summary[
            "internal_invalid_va_count"
        ] == 0
        and
        summary[
            "bank_invalid_va_count"
        ] == 0
    ):

        st.success(
            "🟢 RECONCILIATION HEALTHY"
        )

    elif (
        summary[
            "transaction_match_rate"
        ] >= 98
        and
        summary[
            "amount_match_rate"
        ] >= 98
    ):

        st.warning(
            "🟡 RECONCILIATION WARNING"
        )

    else:

        st.error(
            "🔴 RECONCILIATION CRITICAL"
        )


    # ========================================================
    # DATA QUALITY
    # ========================================================

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
            "### 💰 Nominal Invalid VA"
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


    st.divider()


    # ========================================================
    # ISSUE FMSS & BANK
    # ========================================================

    col1, col2 = st.columns(2)


    # ========================================================
    # ISSUE FMSS
    # ========================================================

    with col1:

        st.markdown(
            "#### 🚨 Issue FMSS"
        )

        if not df_selisih_int.empty:

            display_cols = []

            if "VA_CODE" in df_selisih_int.columns:
                display_cols.append(
                    "VA_CODE"
                )

            if "INT__VA_TYPE" in df_selisih_int.columns:
                display_cols.append(
                    "INT__VA_TYPE"
                )

            if "INT__nominal" in df_selisih_int.columns:
                display_cols.append(
                    "INT__nominal"
                )

            if "ISSUE_TYPE" in df_selisih_int.columns:
                display_cols.append(
                    "ISSUE_TYPE"
                )

            display_int = (
                df_selisih_int[
                    display_cols
                ]
                .rename(
                    columns={
                        "VA_CODE":
                            "KODE VA",

                        "INT__VA_TYPE":
                            "JENIS VA",

                        "INT__nominal":
                            "NOMINAL",

                        "ISSUE_TYPE":
                            "ISSUE"
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

            bank_amount_col = (
                "BNK__"
                +
                config[
                    "bank_credit_col"
                ]
            )

            display_cols = [
                "VA_CODE"
            ]

            if (
                "BNK__VA_TYPE"
                in df_selisih_bnk.columns
            ):

                display_cols.append(
                    "BNK__VA_TYPE"
                )

            if (
                bank_amount_col
                in df_selisih_bnk.columns
            ):

                display_cols.append(
                    bank_amount_col
                )

            display_cols.append(
                "ISSUE_TYPE"
            )

            display_bank = (
                df_selisih_bnk[
                    display_cols
                ]
                .rename(
                    columns={
                        "VA_CODE":
                            "KODE VA",

                        "BNK__VA_TYPE":
                            "JENIS VA",

                        bank_amount_col:
                            "NOMINAL",

                        "ISSUE_TYPE":
                            "ISSUE"
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
        "⚠️ Transaksi dengan VA "
        "Tidak Teridentifikasi"
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
    # EXPORT EXCEL
    # ========================================================

    st.divider()

    st.markdown(
        "### 📥 Download Laporan"
    )

    # --------------------------------------------------------
    # FMSS ISSUE
    # --------------------------------------------------------

    issue_fmss_export = (
        df_selisih_int
        .drop(
            columns=[
                "NOMINAL_MATCH",
                "_occ",
                "_merge"
            ],
            errors="ignore"
        )
        .copy()
    )

    issue_fmss_export.columns = [
        str(c).replace(
            "INT__",
            ""
        )
        for c in issue_fmss_export.columns
    ]


    # --------------------------------------------------------
    # BANK ISSUE
    # --------------------------------------------------------

    issue_bank_export = (
        df_selisih_bnk
        .drop(
            columns=[
                "NOMINAL_MATCH",
                "_occ",
                "_merge"
            ],
            errors="ignore"
        )
        .copy()
    )


    # --------------------------------------------------------
    # MATCHED
    # --------------------------------------------------------

    matched_export = (
        df_matched
        .drop(
            columns=[
                "NOMINAL_MATCH",
                "_occ",
                "_merge"
            ],
            errors="ignore"
        )
        .copy()
    )


    # ========================================================
    # EXCEL WRITER
    # ========================================================

    output = io.BytesIO()

    try:

        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
        ) as writer:

            workbook = writer.book

            header_format = (
                workbook.add_format({
                    "bold": True,
                    "border": 1,
                    "align": "center",
                    "valign": "vcenter"
                })
            )

            money_format = (
                workbook.add_format({
                    "num_format":
                        '#,##0'
                })
            )

            percent_format = (
                workbook.add_format({
                    "num_format":
                        '0.00%'
                })
            )


            # =================================================
            # SUMMARY SHEET
            # =================================================

            summary_df = pd.DataFrame([
                {

                    "RECONCILIATION_ID":
                        summary[
                            "reconciliation_id"
                        ],

                    "BANK":
                        summary[
                            "bank"
                        ],

                    "PROCESSING_TIME":
                        summary[
                            "processing_time"
                        ],

                    "EXPECTED_AMOUNT":
                        summary[
                            "expected_amount"
                        ],

                    "ACTUAL_BANK_AMOUNT":
                        summary[
                            "actual_amount"
                        ],

                    "MATCHED_AMOUNT":
                        summary[
                            "matched_amount"
                        ],

                    "AMOUNT_DIFFERENCE":
                        summary[
                            "amount_difference"
                        ],

                    "TRANSACTION_MATCH_RATE":
                        summary[
                            "transaction_match_rate"
                        ] / 100,

                    "AMOUNT_MATCH_RATE":
                        summary[
                            "amount_match_rate"
                        ] / 100,

                    "FMSS_ISSUE_COUNT":
                        summary[
                            "issue_fmss_count"
                        ],

                    "FMSS_ISSUE_AMOUNT":
                        summary[
                            "issue_fmss_amount"
                        ],

                    "BANK_ISSUE_COUNT":
                        summary[
                            "issue_bank_count"
                        ],

                    "BANK_ISSUE_AMOUNT":
                        summary[
                            "issue_bank_amount"
                        ],

                    "FMSS_INVALID_VA":
                        summary[
                            "internal_invalid_va_count"
                        ],

                    "BANK_INVALID_VA":
                        summary[
                            "bank_invalid_va_count"
                        ],

                    "CARRY_OVER_COUNT":
                        summary[
                            "carry_over_count"
                        ],

                    "FEE_ADJUSTMENT":
                        summary[
                            "fee_adjustment"
                        ],

                    "FEE_REASON":
                        summary[
                            "fee_reason"
                        ]
                }
            ])


            summary_df.to_excel(
                writer,
                sheet_name="SUMMARY",
                index=False
            )

            ws = writer.sheets[
                "SUMMARY"
            ]

            for col_num, value in enumerate(
                summary_df.columns
            ):

                ws.write(
                    0,
                    col_num,
                    value,
                    header_format
                )

            ws.freeze_panes(
                1,
                0
            )

            ws.set_column(
                0,
                len(
                    summary_df.columns
                ) - 1,
                20
            )

            for col_name in [
                "EXPECTED_AMOUNT",
                "ACTUAL_BANK_AMOUNT",
                "MATCHED_AMOUNT",
                "AMOUNT_DIFFERENCE",
                "FMSS_ISSUE_AMOUNT",
                "BANK_ISSUE_AMOUNT",
                "FEE_ADJUSTMENT"
            ]:

                idx = (
                    summary_df.columns
                    .get_loc(
                        col_name
                    )
                )

                ws.set_column(
                    idx,
                    idx,
                    20,
                    money_format
                )

            for col_name in [
                "TRANSACTION_MATCH_RATE",
                "AMOUNT_MATCH_RATE"
            ]:

                idx = (
                    summary_df.columns
                    .get_loc(
                        col_name
                    )
                )

                ws.set_column(
                    idx,
                    idx,
                    22,
                    percent_format
                )


            # =================================================
            # ISSUE FMSS
            # =================================================

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

            ws = writer.sheets[
                "ISSUE_FMSS"
            ]

            ws.freeze_panes(
                1,
                0
            )


            # =================================================
            # ISSUE BANK
            # =================================================

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

            ws = writer.sheets[
                "ISSUE_BANK"
            ]

            ws.freeze_panes(
                1,
                0
            )


            # =================================================
            # MATCHED
            # =================================================

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

            ws = writer.sheets[
                "MATCHED_OK"
            ]

            ws.freeze_panes(
                1,
                0
            )


            # =================================================
            # INVALID FMSS
            # =================================================

            if not df_invalid_int.empty:

                df_invalid_int.to_excel(
                    writer,
                    sheet_name="INVALID_VA_FMSS",
                    index=False
                )

                writer.sheets[
                    "INVALID_VA_FMSS"
                ].freeze_panes(
                    1,
                    0
                )


            # =================================================
            # INVALID BANK
            # =================================================

            if not df_invalid_bnk.empty:

                df_invalid_bnk.to_excel(
                    writer,
                    sheet_name="INVALID_VA_BANK",
                    index=False
                )

                writer.sheets[
                    "INVALID_VA_BANK"
                ].freeze_panes(
                    1,
                    0
                )


        # ====================================================
        # FINALIZE EXCEL BYTES
        # ====================================================

        excel_data = (
            output.getvalue()
        )


        # ====================================================
        # DOWNLOAD
        # ====================================================

        st.download_button(

            label=(
                "📥 Download Laporan "
                "Rekonsiliasi Lengkap (.xlsx)"
            ),

            data=excel_data,

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

        st.success(
            "✅ Laporan Excel berhasil dibuat."
        )

        st.caption(
            "ISSUE_FMSS dapat digunakan sebagai "
            "carry-over untuk rekonsiliasi berikutnya."
        )


    except Exception as e:

        st.error(
            "❌ Data berhasil diproses, "
            "tetapi laporan Excel gagal dibuat."
        )

        if mode_dev:

            st.code(
                traceback.format_exc()
            )

        else:

            st.caption(
                f"Technical detail: {e}"
            )


# ============================================================
# EMPTY STATE
# ============================================================

elif (
    pilihan_bank == ""
    and (
        file_int
        or file_bnk
    )
):

    st.info(
        "💡 Silakan pilih Bank Sumber Mutasi "
        "terlebih dahulu."
    )
