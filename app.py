import streamlit as st
import pandas as pd
import io
import re
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
    "Dashboard rekonsiliasi otomatis antara data deposit FMSS "
    "dengan mutasi BRIVA Fastpay dan BRIVA Rajabiller."
)

st.divider()


# ============================================================
# KONFIGURASI PREFIX
# ============================================================

PREFIX_FASTPAY = "57888"
PREFIX_RAJABILLER = "57708"

TYPE_FASTPAY = "BRIVA FASTPAY"
TYPE_RAJABILLER = "BRIVA RAJABILLER"


# ============================================================
# KONFIGURASI FEE
# ============================================================
#
# Default:
# FMSS nominal + Rp1.000 = nominal bank
#
# Jika ternyata Rajabiller memiliki fee berbeda,
# tinggal ubah angka di UI.
#

DEFAULT_FEE_FASTPAY = 1000
DEFAULT_FEE_RAJABILLER = 1000


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_STATE = {

    "processed": False,

    "result_fastpay": {},
    "result_rajabiller": {},

    "invalid_fmss": pd.DataFrame(),
    "invalid_bank_fastpay": pd.DataFrame(),
    "invalid_bank_rajabiller": pd.DataFrame(),

    "summary": {},

    "last_process_id": "",

}


for key, value in DEFAULT_STATE.items():

    if key not in st.session_state:

        st.session_state[key] = value


# ============================================================
# HELPER - READ FILE
# ============================================================

def read_file(uploaded_file):

    uploaded_file.seek(0)

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):

        return pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python"
        )

    elif filename.endswith(".xlsx"):

        return pd.read_excel(
            uploaded_file
        )

    else:

        raise ValueError(
            f"Format file tidak didukung: {filename}"
        )


# ============================================================
# HELPER - VALIDATE COLUMN
# ============================================================

def validate_columns(
    df,
    required_columns,
    source_name
):

    missing = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing:

        raise ValueError(
            f"{source_name} tidak memiliki kolom wajib: "
            f"{', '.join(missing)}"
        )


# ============================================================
# HELPER - PARSE NOMINAL
# ============================================================

def parse_nominal(series):

    if pd.api.types.is_numeric_dtype(series):

        return pd.to_numeric(
            series,
            errors="coerce"
        ).fillna(0)

    text = (
        series
        .astype(str)
        .str.strip()
    )

    # Format Indonesia:
    # 1.999.000
    # 1.999.000,00

    text = (
        text
        .str.replace(
            "Rp",
            "",
            regex=False
        )
        .str.replace(
            "rp",
            "",
            regex=False
        )
        .str.replace(
            " ",
            "",
            regex=False
        )
    )

    text = (
        text
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
    )

    return pd.to_numeric(
        text,
        errors="coerce"
    ).fillna(0)


# ============================================================
# HELPER - EXTRACT VA
# ============================================================

def extract_va(
    text,
    prefixes
):

    if pd.isna(text):

        return None, None

    text = str(text)

    for prefix, va_type in prefixes.items():

        # Prefix + 5 sampai 15 digit
        pattern = (
            rf"({re.escape(prefix)}\d{{5,15}})"
        )

        match = re.search(
            pattern,
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


# ============================================================
# HELPER - APPLY VA
# ============================================================

def apply_va(
    df,
    source_column,
    prefixes
):

    df = df.copy()

    extracted = (
        df[source_column]
        .apply(
            lambda x:
            extract_va(
                x,
                prefixes
            )
        )
    )

    df["VA_CODE"] = extracted.apply(
        lambda x: x[0]
    )

    df["VA_TYPE"] = extracted.apply(
        lambda x: x[1]
    )

    return df


# ============================================================
# HELPER - OCCURRENCE KEY
# ============================================================
#
# Penting untuk transaksi duplicate.
#
# Contoh:
#
# VA 578881234
# Rp 100.000
#
# muncul 3x.
#
# Maka transaksi ke-1, ke-2, ke-3
# tetap dapat dicocokkan 1-to-1.
#

def add_occurrence_key(df):

    df = df.copy()

    df["_OCCURRENCE"] = (
        df
        .groupby(
            [
                "VA_CODE",
                "NOMINAL_MATCH"
            ],
            dropna=False
        )
        .cumcount()
    )

    return df


# ============================================================
# RECONCILIATION ENGINE
# ============================================================

def reconcile_pair(
    df_internal,
    df_bank,
    fee_adjustment,
    recon_name,
    recon_id
):

    # --------------------------------------------------------
    # INTERNAL
    # --------------------------------------------------------

    validate_columns(
        df_internal,
        [
            "status",
            "keterangan",
            "nominal"
        ],
        f"FMSS {recon_name}"
    )

    # --------------------------------------------------------
    # BANK
    # --------------------------------------------------------

    validate_columns(
        df_bank,
        [
            "DESK_TRAN",
            "MUTASI_KREDIT"
        ],
        f"Bank {recon_name}"
    )

    # --------------------------------------------------------
    # FMSS SUCCESS
    # --------------------------------------------------------

    df_int = (
        df_internal[
            df_internal[
                "status"
            ]
            .astype(str)
            .str.upper()
            .eq("SUKSES")
        ]
        .copy()
    )

    total_fmss_success = len(df_int)

    # --------------------------------------------------------
    # EXTRACT VA FMSS
    # --------------------------------------------------------

    df_int = apply_va(
        df_int,
        "keterangan",
        {
            PREFIX_FASTPAY:
                TYPE_FASTPAY,

            PREFIX_RAJABILLER:
                TYPE_RAJABILLER
        }
    )

    # --------------------------------------------------------
    # NOMINAL FMSS
    # --------------------------------------------------------

    df_int[
        "NOMINAL_ORIGINAL"
    ] = parse_nominal(
        df_int[
            "nominal"
        ]
    )

    df_int[
        "NOMINAL_MATCH"
    ] = (
        df_int[
            "NOMINAL_ORIGINAL"
        ]
        + fee_adjustment
    )

    # --------------------------------------------------------
    # INVALID FMSS
    # --------------------------------------------------------

    invalid_fmss = (
        df_int[
            df_int[
                "VA_CODE"
            ].isna()
        ]
        .copy()
    )

    # --------------------------------------------------------
    # VALID FMSS
    # --------------------------------------------------------

    valid_fmss = (
        df_int[
            df_int[
                "VA_CODE"
            ].notna()
        ]
        .copy()
        .reset_index(
            drop=True
        )
    )

    # --------------------------------------------------------
    # BANK CREDIT
    # --------------------------------------------------------

    df_bank = df_bank.copy()

    df_bank[
        "MUTASI_KREDIT_NUM"
    ] = parse_nominal(
        df_bank[
            "MUTASI_KREDIT"
        ]
    )

    df_bank = (
        df_bank[
            df_bank[
                "MUTASI_KREDIT_NUM"
            ] > 0
        ]
        .copy()
    )

    total_bank_credit = len(df_bank)

    # --------------------------------------------------------
    # EXTRACT BANK VA
    # --------------------------------------------------------

    bank_prefix = (
        PREFIX_FASTPAY
        if recon_name == "FASTPAY"
        else PREFIX_RAJABILLER
    )

    bank_type = (
        TYPE_FASTPAY
        if recon_name == "FASTPAY"
        else TYPE_RAJABILLER
    )

    df_bank = apply_va(
        df_bank,
        "DESK_TRAN",
        {
            bank_prefix:
                bank_type
        }
    )

    # --------------------------------------------------------
    # INVALID BANK
    # --------------------------------------------------------

    invalid_bank = (
        df_bank[
            df_bank[
                "VA_CODE"
            ].isna()
        ]
        .copy()
    )

    # --------------------------------------------------------
    # VALID BANK
    # --------------------------------------------------------

    valid_bank = (
        df_bank[
            df_bank[
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
            "MUTASI_KREDIT_NUM"
        ]
    )

    # --------------------------------------------------------
    # ADD OCCURRENCE
    # --------------------------------------------------------

    valid_fmss = add_occurrence_key(
        valid_fmss
    )

    valid_bank = add_occurrence_key(
        valid_bank
    )

    # ========================================================
    # MERGE 1-TO-1
    # ========================================================

    matched = pd.merge(
        valid_fmss,
        valid_bank,
        on=[
            "VA_CODE",
            "NOMINAL_MATCH",
            "_OCCURRENCE"
        ],
        how="outer",
        indicator=True,
        suffixes=(
            "_FMSS",
            "_BANK"
        )
    )

    # --------------------------------------------------------
    # MATCHED
    # --------------------------------------------------------

    matched_ok = (
        matched[
            matched[
                "_merge"
            ].eq("both")
        ]
        .copy()
    )

    matched_ok[
        "MATCH_STATUS"
    ] = "MATCHED"

    matched_ok[
        "RECONCILIATION_ID"
    ] = recon_id

    matched_ok[
        "RECON_TYPE"
    ] = recon_name

    # --------------------------------------------------------
    # FMSS ONLY
    # --------------------------------------------------------

    fmss_only = (
        matched[
            matched[
                "_merge"
            ].eq("left_only")
        ]
        .copy()
    )

    fmss_only[
        "ISSUE_TYPE"
    ] = "FMSS_ONLY"

    fmss_only[
        "RECONCILIATION_ID"
    ] = recon_id

    fmss_only[
        "RECON_TYPE"
    ] = recon_name

    # --------------------------------------------------------
    # BANK ONLY
    # --------------------------------------------------------

    bank_only = (
        matched[
            matched[
                "_merge"
            ].eq("right_only")
        ]
        .copy()
    )

    bank_only[
        "ISSUE_TYPE"
    ] = "BANK_ONLY"

    bank_only[
        "RECONCILIATION_ID"
    ] = recon_id

    bank_only[
        "RECON_TYPE"
    ] = recon_name

    # ========================================================
    # AMOUNT MISMATCH DETECTION
    # ========================================================
    #
    # Jika VA sama tetapi nominal berbeda,
    # jangan hanya dianggap FMSS_ONLY / BANK_ONLY.
    #
    # Kita tandai sebagai AMOUNT_MISMATCH.
    #

    fmss_only_va = set(
        fmss_only[
            "VA_CODE"
        ]
        .dropna()
    )

    bank_only_va = set(
        bank_only[
            "VA_CODE"
        ]
        .dropna()
    )

    common_issue_va = (
        fmss_only_va
        &
        bank_only_va
    )

    if common_issue_va:

        fmss_only.loc[
            fmss_only[
                "VA_CODE"
            ].isin(
                common_issue_va
            ),
            "ISSUE_TYPE"
        ] = "AMOUNT_MISMATCH"

        bank_only.loc[
            bank_only[
                "VA_CODE"
            ].isin(
                common_issue_va
            ),
            "ISSUE_TYPE"
        ] = "AMOUNT_MISMATCH"

    # ========================================================
    # SUMMARY
    # ========================================================

    expected_amount = (
        valid_fmss[
            "NOMINAL_MATCH"
        ].sum()
    )

    bank_amount = (
        valid_bank[
            "NOMINAL_MATCH"
        ].sum()
    )

    matched_amount = (
        matched_ok[
            "NOMINAL_MATCH"
        ].sum()
        if not matched_ok.empty
        else 0
    )

    fmss_issue_amount = (
        fmss_only[
            "NOMINAL_MATCH"
        ].sum()
        if not fmss_only.empty
        else 0
    )

    bank_issue_amount = (
        bank_only[
            "NOMINAL_MATCH"
        ].sum()
        if not bank_only.empty
        else 0
    )

    valid_fmss_count = len(
        valid_fmss
    )

    matched_count = len(
        matched_ok
    )

    transaction_match_rate = (
        matched_count
        /
        valid_fmss_count
        *
        100
        if valid_fmss_count > 0
        else 0
    )

    amount_match_rate = (
        matched_amount
        /
        expected_amount
        *
        100
        if expected_amount > 0
        else 0
    )

    amount_difference = (
        bank_amount
        -
        expected_amount
    )

    summary = {

        "recon_type":
            recon_name,

        "reconciliation_id":
            recon_id,

        "fmss_success":
            total_fmss_success,

        "fmss_valid_va":
            valid_fmss_count,

        "fmss_invalid_va":
            len(invalid_fmss),

        "bank_credit":
            total_bank_credit,

        "bank_valid_va":
            len(valid_bank),

        "bank_invalid_va":
            len(invalid_bank),

        "matched":
            matched_count,

        "issue_fmss":
            len(fmss_only),

        "issue_bank":
            len(bank_only),

        "expected_amount":
            expected_amount,

        "bank_amount":
            bank_amount,

        "matched_amount":
            matched_amount,

        "fmss_issue_amount":
            fmss_issue_amount,

        "bank_issue_amount":
            bank_issue_amount,

        "amount_difference":
            amount_difference,

        "transaction_match_rate":
            transaction_match_rate,

        "amount_match_rate":
            amount_match_rate,

        "fee_adjustment":
            fee_adjustment

    }

    return {

        "matched":
            matched_ok,

        "issue_fmss":
            fmss_only,

        "issue_bank":
            bank_only,

        "invalid_fmss":
            invalid_fmss,

        "invalid_bank":
            invalid_bank,

        "summary":
            summary

    }


# ============================================================
# DISPLAY MONEY
# ============================================================

def rupiah(value):

    try:

        return (
            f"Rp {float(value):,.0f}"
        )

    except:

        return "Rp 0"


# ============================================================
# DISPLAY SUMMARY
# ============================================================

def render_recon_summary(
    result,
    title
):

    summary = result[
        "summary"
    ]

    st.markdown(
        f"### {title}"
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Matched",
        f"{summary['matched']:,}"
    )

    c2.metric(
        "Issue FMSS",
        f"{summary['issue_fmss']:,}"
    )

    c3.metric(
        "Issue Bank",
        f"{summary['issue_bank']:,}"
    )

    c4.metric(
        "Match Rate",
        f"{summary['transaction_match_rate']:.2f}%"
    )

    c5, c6, c7, c8 = st.columns(4)

    c5.metric(
        "Expected",
        rupiah(
            summary[
                "expected_amount"
            ]
        )
    )

    c6.metric(
        "Actual Bank",
        rupiah(
            summary[
                "bank_amount"
            ]
        )
    )

    c7.metric(
        "Matched Amount",
        rupiah(
            summary[
                "matched_amount"
            ]
        )
    )

    c8.metric(
        "Amount Difference",
        rupiah(
            summary[
                "amount_difference"
            ]
        )
    )

    if (
        summary[
            "transaction_match_rate"
        ] >= 99.5
        and
        summary[
            "amount_match_rate"
        ] >= 99.5
    ):

        st.success(
            "🟢 Rekonsiliasi sehat"
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
            "🟡 Rekonsiliasi perlu review"
        )

    else:

        st.error(
            "🔴 Rekonsiliasi memiliki issue signifikan"
        )


# ============================================================
# UI - UPLOAD
# ============================================================

st.subheader(
    "1. Upload Data"
)

col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        "**📄 FMSS**"
    )

    file_fmss = st.file_uploader(
        "Upload data FMSS",
        type=[
            "csv",
            "xlsx"
        ],
        key="file_fmss"
    )


with col2:

    st.markdown(
        "**🏦 BRIVA Fastpay — 57888**"
    )

    file_fastpay = st.file_uploader(
        "Upload mutasi BRIVA 57888",
        type=[
            "csv",
            "xlsx"
        ],
        key="file_fastpay"
    )


with col3:

    st.markdown(
        "**🏦 BRIVA Rajabiller — 57708**"
    )

    file_rajabiller = st.file_uploader(
        "Upload mutasi BRIVA 57708",
        type=[
            "csv",
            "xlsx"
        ],
        key="file_rajabiller"
    )


# ============================================================
# FEE CONFIGURATION
# ============================================================

st.subheader(
    "2. Konfigurasi Fee"
)

fee1, fee2 = st.columns(2)


with fee1:

    fee_fastpay = st.number_input(
        "Fee Fastpay (57888)",
        min_value=0,
        value=DEFAULT_FEE_FASTPAY,
        step=500
    )


with fee2:

    fee_rajabiller = st.number_input(
        "Fee Rajabiller (57708)",
        min_value=0,
        value=DEFAULT_FEE_RAJABILLER,
        step=500
    )


st.caption(
    "Rumus pencocokan: Nominal FMSS + Fee = Nominal mutasi bank."
)


# ============================================================
# PROCESS BUTTON
# ============================================================

ready = (
    file_fmss
    and file_fastpay
    and file_rajabiller
)


if ready:

    st.divider()

    if st.button(
        "🚀 MULAI REKONSILIASI",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "Memproses rekonsiliasi Fastpay dan Rajabiller..."
            ):

                # ------------------------------------------------
                # READ FILE
                # ------------------------------------------------

                df_fmss = read_file(
                    file_fmss
                )

                df_fastpay = read_file(
                    file_fastpay
                )

                df_rajabiller = read_file(
                    file_rajabiller
                )

                # ------------------------------------------------
                # PROCESS ID
                # ------------------------------------------------

                process_id = datetime.now().strftime(
                    "REC-%Y%m%d-%H%M%S"
                )

                # ------------------------------------------------
                # RECON FASTPAY
                # ------------------------------------------------

                result_fastpay = reconcile_pair(

                    df_fmss,

                    df_fastpay,

                    fee_fastpay,

                    "FASTPAY",

                    process_id
                )

                # ------------------------------------------------
                # RECON RAJABILLER
                # ------------------------------------------------

                result_rajabiller = reconcile_pair(

                    df_fmss,

                    df_rajabiller,

                    fee_rajabiller,

                    "RAJABILLER",

                    process_id
                )

                # ------------------------------------------------
                # INVALID FMSS
                #
                # Ambil dari seluruh FMSS.
                #
                # Yang valid:
                # 57888 atau 57708
                #
                # Yang tidak punya prefix:
                # INVALID
                # ------------------------------------------------

                validate_columns(
                    df_fmss,
                    [
                        "status",
                        "keterangan",
                        "nominal"
                    ],
                    "FMSS"
                )

                df_fmss_success = (
                    df_fmss[
                        df_fmss[
                            "status"
                        ]
                        .astype(str)
                        .str.upper()
                        .eq("SUKSES")
                    ]
                    .copy()
                )

                df_fmss_success = apply_va(
                    df_fmss_success,
                    "keterangan",
                    {
                        PREFIX_FASTPAY:
                            TYPE_FASTPAY,

                        PREFIX_RAJABILLER:
                            TYPE_RAJABILLER
                    }
                )

                df_invalid_fmss = (
                    df_fmss_success[
                        df_fmss_success[
                            "VA_CODE"
                        ].isna()
                    ]
                    .copy()
                )

                # ------------------------------------------------
                # INVALID BANK
                # ------------------------------------------------

                df_invalid_bank_fastpay = (
                    result_fastpay[
                        "invalid_bank"
                    ]
                    .copy()
                )

                df_invalid_bank_rajabiller = (
                    result_rajabiller[
                        "invalid_bank"
                    ]
                    .copy()
                )

                # ------------------------------------------------
                # SAVE STATE
                # ------------------------------------------------

                st.session_state[
                    "result_fastpay"
                ] = result_fastpay

                st.session_state[
                    "result_rajabiller"
                ] = result_rajabiller

                st.session_state[
                    "invalid_fmss"
                ] = df_invalid_fmss

                st.session_state[
                    "invalid_bank_fastpay"
                ] = df_invalid_bank_fastpay

                st.session_state[
                    "invalid_bank_rajabiller"
                ] = df_invalid_bank_rajabiller

                st.session_state[
                    "last_process_id"
                ] = process_id

                # ------------------------------------------------
                # TOTAL SUMMARY
                # ------------------------------------------------

                s1 = result_fastpay[
                    "summary"
                ]

                s2 = result_rajabiller[
                    "summary"
                ]

                total_valid_fmss = (
                    s1["fmss_valid_va"]
                    +
                    s2["fmss_valid_va"]
                )

                total_matched = (
                    s1["matched"]
                    +
                    s2["matched"]
                )

                total_expected = (
                    s1["expected_amount"]
                    +
                    s2["expected_amount"]
                )

                total_bank = (
                    s1["bank_amount"]
                    +
                    s2["bank_amount"]
                )

                total_matched_amount = (
                    s1["matched_amount"]
                    +
                    s2["matched_amount"]
                )

                total_issue_fmss = (
                    s1["issue_fmss"]
                    +
                    s2["issue_fmss"]
                )

                total_issue_bank = (
                    s1["issue_bank"]
                    +
                    s2["issue_bank"]
                )

                total_difference = (
                    total_bank
                    -
                    total_expected
                )

                total_match_rate = (
                    total_matched
                    /
                    total_valid_fmss
                    *
                    100
                    if total_valid_fmss
                    else 0
                )

                total_amount_rate = (
                    total_matched_amount
                    /
                    total_expected
                    *
                    100
                    if total_expected
                    else 0
                )

                st.session_state[
                    "summary"
                ] = {

                    "process_id":
                        process_id,

                    "total_valid_fmss":
                        total_valid_fmss,

                    "total_matched":
                        total_matched,

                    "total_issue_fmss":
                        total_issue_fmss,

                    "total_issue_bank":
                        total_issue_bank,

                    "total_expected":
                        total_expected,

                    "total_bank":
                        total_bank,

                    "total_matched_amount":
                        total_matched_amount,

                    "total_difference":
                        total_difference,

                    "total_match_rate":
                        total_match_rate,

                    "total_amount_rate":
                        total_amount_rate

                }

                st.session_state[
                    "processed"
                ] = True

            st.success(
                "✅ Rekonsiliasi selesai."
            )

        except Exception as e:

            st.session_state[
                "processed"
            ] = False

            st.error(
                "❌ Rekonsiliasi gagal diproses."
            )

            with st.expander(
                "Technical Error"
            ):

                st.code(
                    traceback.format_exc()
                )


else:

    missing = []

    if not file_fmss:
        missing.append(
            "FMSS"
        )

    if not file_fastpay:
        missing.append(
            "Mutasi BRIVA 57888"
        )

    if not file_rajabiller:
        missing.append(
            "Mutasi BRIVA 57708"
        )

    st.info(
        "Upload terlebih dahulu: "
        + ", ".join(missing)
    )


# ============================================================
# DISPLAY RESULT
# ============================================================

if st.session_state[
    "processed"
]:

    result_fastpay = (
        st.session_state[
            "result_fastpay"
        ]
    )

    result_rajabiller = (
        st.session_state[
            "result_rajabiller"
        ]
    )

    invalid_fmss = (
        st.session_state[
            "invalid_fmss"
        ]
    )

    invalid_bank_fastpay = (
        st.session_state[
            "invalid_bank_fastpay"
        ]
    )

    invalid_bank_rajabiller = (
        st.session_state[
            "invalid_bank_rajabiller"
        ]
    )

    summary = (
        st.session_state[
            "summary"
        ]
    )

    st.divider()

    # ========================================================
    # TOTAL SUMMARY
    # ========================================================

    st.subheader(
        "🎯 Overall Reconciliation"
    )

    a1, a2, a3, a4 = st.columns(4)

    a1.metric(
        "Matched",
        f"{summary['total_matched']:,}"
    )

    a2.metric(
        "Issue FMSS",
        f"{summary['total_issue_fmss']:,}"
    )

    a3.metric(
        "Issue Bank",
        f"{summary['total_issue_bank']:,}"
    )

    a4.metric(
        "Overall Match Rate",
        f"{summary['total_match_rate']:.2f}%"
    )

    a5, a6, a7, a8 = st.columns(4)

    a5.metric(
        "Expected",
        rupiah(
            summary[
                "total_expected"
            ]
        )
    )

    a6.metric(
        "Actual Bank",
        rupiah(
            summary[
                "total_bank"
            ]
        )
    )

    a7.metric(
        "Matched Amount",
        rupiah(
            summary[
                "total_matched_amount"
            ]
        )
    )

    a8.metric(
        "Amount Difference",
        rupiah(
            summary[
                "total_difference"
            ]
        )
    )

    if (
        summary[
            "total_match_rate"
        ] >= 99.5
        and
        summary[
            "total_amount_rate"
        ] >= 99.5
    ):

        st.success(
            "🟢 Overall reconciliation HEALTHY"
        )

    elif (
        summary[
            "total_match_rate"
        ] >= 98
        and
        summary[
            "total_amount_rate"
        ] >= 98
    ):

        st.warning(
            "🟡 Overall reconciliation WARNING"
        )

    else:

        st.error(
            "🔴 Overall reconciliation CRITICAL"
        )


    # ========================================================
    # FASTPAY
    # ========================================================

    st.divider()

    render_recon_summary(
        result_fastpay,
        "🏦 BRIVA FASTPAY — 57888"
    )


    # ========================================================
    # RAJABILLER
    # ========================================================

    st.divider()

    render_recon_summary(
        result_rajabiller,
        "🏦 BRIVA RAJABILLER — 57708"
    )


    # ========================================================
    # ISSUE DETAIL
    # ========================================================

    st.divider()

    st.subheader(
        "🚨 Detail Issue"
    )

    issue_tab1, issue_tab2 = st.tabs(
        [
            "FMSS Issue",
            "Bank Issue"
        ]
    )


    # ========================================================
    # FMSS ISSUE
    # ========================================================

    with issue_tab1:

        fastpay_issue = (
            result_fastpay[
                "issue_fmss"
            ]
            .copy()
        )

        rajabiller_issue = (
            result_rajabiller[
                "issue_fmss"
            ]
            .copy()
        )

        if not fastpay_issue.empty:

            st.markdown(
                "### 🚨 FMSS Issue — Fastpay 57888"
            )

            display = prepare_issue_fmss(
                fastpay_issue
            )

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue FMSS Fastpay."
            )


        if not rajabiller_issue.empty:

            st.markdown(
                "### 🚨 FMSS Issue — Rajabiller 57708"
            )

            display = prepare_issue_fmss(
                rajabiller_issue
            )

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue FMSS Rajabiller."
            )


    # ========================================================
    # BANK ISSUE
    # ========================================================

    with issue_tab2:

        fastpay_bank_issue = (
            result_fastpay[
                "issue_bank"
            ]
            .copy()
        )

        rajabiller_bank_issue = (
            result_rajabiller[
                "issue_bank"
            ]
            .copy()
        )

        if not fastpay_bank_issue.empty:

            st.markdown(
                "### 🚨 Bank Issue — Fastpay 57888"
            )

            display = prepare_issue_bank(
                fastpay_bank_issue
            )

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue Bank Fastpay."
            )


        if not rajabiller_bank_issue.empty:

            st.markdown(
                "### 🚨 Bank Issue — Rajabiller 57708"
            )

            display = prepare_issue_bank(
                rajabiller_bank_issue
            )

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue Bank Rajabiller."
            )


    # ========================================================
    # INVALID VA
    # ========================================================

    st.divider()

    with st.expander(
        "⚠️ Transaksi dengan VA Tidak Teridentifikasi"
    ):

        inv1, inv2 = st.columns(2)


        with inv1:

            st.markdown(
                "### FMSS Invalid VA"
            )

            if not invalid_fmss.empty:

                st.dataframe(
                    invalid_fmss,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.success(
                    "Tidak ada FMSS invalid VA."
                )


        with inv2:

            st.markdown(
                "### Bank Invalid VA"
            )

            all_invalid_bank = pd.concat(
                [
                    invalid_bank_fastpay.assign(
                        SUMBER="BRIVA FASTPAY"
                    ),
                    invalid_bank_rajabiller.assign(
                        SUMBER="BRIVA RAJABILLER"
                    )
                ],
                ignore_index=True
            )

            if not all_invalid_bank.empty:

                st.dataframe(
                    all_invalid_bank,
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

    st.subheader(
        "📥 Download Laporan"
    )

    try:

        output = io.BytesIO()

        # ====================================================
        # XlsxWriter
        # ====================================================

        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
        ) as writer:

            workbook = writer.book

            header_format = workbook.add_format({
                "bold": True,
                "border": 1,
                "align": "center",
                "valign": "vcenter"
            })

            money_format = workbook.add_format({
                "num_format": '#,##0'
            })

            percent_format = workbook.add_format({
                "num_format": '0.00%'
            })

            # ------------------------------------------------
            # SUMMARY
            # ------------------------------------------------

            overall_df = pd.DataFrame([
                {

                    "PROCESS_ID":
                        summary[
                            "process_id"
                        ],

                    "TOTAL_MATCHED":
                        summary[
                            "total_matched"
                        ],

                    "TOTAL_ISSUE_FMSS":
                        summary[
                            "total_issue_fmss"
                        ],

                    "TOTAL_ISSUE_BANK":
                        summary[
                            "total_issue_bank"
                        ],

                    "TOTAL_EXPECTED":
                        summary[
                            "total_expected"
                        ],

                    "TOTAL_BANK":
                        summary[
                            "total_bank"
                        ],

                    "TOTAL_MATCHED_AMOUNT":
                        summary[
                            "total_matched_amount"
                        ],

                    "TOTAL_AMOUNT_DIFFERENCE":
                        summary[
                            "total_difference"
                        ],

                    "OVERALL_MATCH_RATE":
                        summary[
                            "total_match_rate"
                        ] / 100,

                    "OVERALL_AMOUNT_RATE":
                        summary[
                            "total_amount_rate"
                        ] / 100

                }
            ])

            overall_df.to_excel(
                writer,
                sheet_name="SUMMARY",
                index=False
            )


            # ------------------------------------------------
            # SUMMARY FASTPAY
            # ------------------------------------------------

            summary_fastpay_df = pd.DataFrame([
                result_fastpay[
                    "summary"
                ]
            ])

            summary_fastpay_df.to_excel(
                writer,
                sheet_name="SUMMARY_FASTPAY",
                index=False
            )


            # ------------------------------------------------
            # SUMMARY RAJABILLER
            # ------------------------------------------------

            summary_rajabiller_df = pd.DataFrame([
                result_rajabiller[
                    "summary"
                ]
            ])

            summary_rajabiller_df.to_excel(
                writer,
                sheet_name="SUMMARY_RAJABILLER",
                index=False
            )


            # ------------------------------------------------
            # MATCHED FASTPAY
            # ------------------------------------------------

            matched_fastpay = (
                result_fastpay[
                    "matched"
                ]
                .copy()
            )

            if matched_fastpay.empty:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada matched Fastpay."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="MATCHED_FASTPAY",
                    index=False
                )

            else:

                matched_fastpay.to_excel(
                    writer,
                    sheet_name="MATCHED_FASTPAY",
                    index=False
                )


            # ------------------------------------------------
            # MATCHED RAJABILLER
            # ------------------------------------------------

            matched_rajabiller = (
                result_rajabiller[
                    "matched"
                ]
                .copy()
            )

            if matched_rajabiller.empty:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada matched Rajabiller."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="MATCHED_RAJABILLER",
                    index=False
                )

            else:

                matched_rajabiller.to_excel(
                    writer,
                    sheet_name="MATCHED_RAJABILLER",
                    index=False
                )


            # ------------------------------------------------
            # ISSUE FMSS
            # ------------------------------------------------

            issue_fmss_export = pd.concat(
                [
                    result_fastpay[
                        "issue_fmss"
                    ].assign(
                        SUMBER="BRIVA FASTPAY"
                    ),

                    result_rajabiller[
                        "issue_fmss"
                    ].assign(
                        SUMBER="BRIVA RAJABILLER"
                    )
                ],
                ignore_index=True
            )

            if issue_fmss_export.empty:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada issue FMSS."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="ISSUE_FMSS",
                    index=False
                )

            else:

                issue_fmss_export.to_excel(
                    writer,
                    sheet_name="ISSUE_FMSS",
                    index=False
                )


            # ------------------------------------------------
            # ISSUE BANK
            # ------------------------------------------------

            issue_bank_export = pd.concat(
                [
                    result_fastpay[
                        "issue_bank"
                    ].assign(
                        SUMBER="BRIVA FASTPAY"
                    ),

                    result_rajabiller[
                        "issue_bank"
                    ].assign(
                        SUMBER="BRIVA RAJABILLER"
                    )
                ],
                ignore_index=True
            )

            if issue_bank_export.empty:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada issue Bank."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="ISSUE_BANK",
                    index=False
                )

            else:

                issue_bank_export.to_excel(
                    writer,
                    sheet_name="ISSUE_BANK",
                    index=False
                )


            # ------------------------------------------------
            # INVALID FMSS
            # ------------------------------------------------

            if invalid_fmss.empty:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada FMSS invalid VA."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="INVALID_VA_FMSS",
                    index=False
                )

            else:

                invalid_fmss.to_excel(
                    writer,
                    sheet_name="INVALID_VA_FMSS",
                    index=False
                )


            # ------------------------------------------------
            # INVALID BANK
            # ------------------------------------------------

            all_invalid_bank_export = pd.concat(
                [
                    invalid_bank_fastpay.assign(
                        SUMBER="BRIVA FASTPAY"
                    ),

                    invalid_bank_rajabiller.assign(
                        SUMBER="BRIVA RAJABILLER"
                    )
                ],
                ignore_index=True
            )

            if all_invalid_bank_export.empty:

                pd.DataFrame({
                    "INFO": [
                        "Tidak ada Bank invalid VA."
                    ]
                }).to_excel(
                    writer,
                    sheet_name="INVALID_VA_BANK",
                    index=False
                )

            else:

                all_invalid_bank_export.to_excel(
                    writer,
                    sheet_name="INVALID_VA_BANK",
                    index=False
                )


            # ------------------------------------------------
            # FORMAT ALL SHEETS
            # ------------------------------------------------

            for sheet_name in writer.sheets:

                worksheet = writer.sheets[
                    sheet_name
                ]

                worksheet.freeze_panes(
                    1,
                    0
                )

                worksheet.autofilter(
                    0,
                    0,
                    0,
                    20
                )

                worksheet.set_column(
                    0,
                    30,
                    18
                )

                # Header
                worksheet.set_row(
                    0,
                    22
                )


            # ------------------------------------------------
            # SPECIAL SUMMARY FORMAT
            # ------------------------------------------------

            ws = writer.sheets[
                "SUMMARY"
            ]

            for col_num, col_name in enumerate(
                overall_df.columns
            ):

                ws.write(
                    0,
                    col_num,
                    col_name,
                    header_format
                )

            # Money
            for col_name in [
                "TOTAL_EXPECTED",
                "TOTAL_BANK",
                "TOTAL_MATCHED_AMOUNT",
                "TOTAL_AMOUNT_DIFFERENCE"
            ]:

                idx = (
                    overall_df.columns
                    .get_loc(
                        col_name
                    )
                )

                ws.set_column(
                    idx,
                    idx,
                    22,
                    money_format
                )

            # Percentage
            for col_name in [
                "OVERALL_MATCH_RATE",
                "OVERALL_AMOUNT_RATE"
            ]:

                idx = (
                    overall_df.columns
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


        excel_data = output.getvalue()

        st.download_button(

            label=(
                "📥 Download Laporan Rekonsiliasi Lengkap (.xlsx)"
            ),

            data=excel_data,

            file_name=(
                f"Laporan_Rekonsiliasi_"
                f"{summary['process_id']}.xlsx"
            ),

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            type="primary",

            use_container_width=True
        )

        st.caption(
            f"Process ID: {summary['process_id']}"
        )

    except Exception as e:

        st.error(
            "❌ Rekonsiliasi berhasil, tetapi "
            "file Excel gagal dibuat."
        )

        st.code(
            traceback.format_exc()
        )


# ============================================================
# HELPER DISPLAY ISSUE FMSS
# ============================================================

def prepare_issue_fmss(df):

    if df.empty:

        return df

    result = pd.DataFrame()

    if "VA_CODE" in df.columns:

        result["KODE VA"] = (
            df["VA_CODE"]
        )

    if "VA_TYPE" in df.columns:

        result["JENIS VA"] = (
            df["VA_TYPE"]
        )

    if "NOMINAL_ORIGINAL" in df.columns:

        result["NOMINAL FMSS"] = (
            df["NOMINAL_ORIGINAL"]
        )

    if "NOMINAL_MATCH" in df.columns:

        result["NOMINAL BANK EXPECTED"] = (
            df["NOMINAL_MATCH"]
        )

    if "ISSUE_TYPE" in df.columns:

        result["ISSUE"] = (
            df["ISSUE_TYPE"]
        )

    return result


# ============================================================
# HELPER DISPLAY ISSUE BANK
# ============================================================

def prepare_issue_bank(df):

    if df.empty:

        return df

    result = pd.DataFrame()

    if "VA_CODE" in df.columns:

        result["KODE VA"] = (
            df["VA_CODE"]
        )

    if "VA_TYPE" in df.columns:

        result["JENIS VA"] = (
            df["VA_TYPE"]
        )

    if "MUTASI_KREDIT_NUM" in df.columns:

        result["NOMINAL BANK"] = (
            df["MUTASI_KREDIT_NUM"]
        )

    if "ISSUE_TYPE" in df.columns:

        result["ISSUE"] = (
            df["ISSUE_TYPE"]
        )

    return result
