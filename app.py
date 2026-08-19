import streamlit as st
import pandas as pd
import re
import io
from collections import defaultdict

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Rekonsiliasi Bank Fastpay",
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

defaults = {
    "sudah_diproses": False,
    "df_matched": pd.DataFrame(),
    "df_selisih_int": pd.DataFrame(),
    "df_selisih_bnk": pd.DataFrame(),
    "df_invalid_int": pd.DataFrame(),
    "df_invalid_bnk": pd.DataFrame(),
    "pilihan_bank_terakhir": "",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# HELPER
# ============================================================

def read_file(uploaded_file):
    """
    Membaca CSV/XLSX dengan seluruh kolom sebagai object/string
    agar VA tidak berubah menjadi angka.
    """

    uploaded_file.seek(0)

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        return pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python",
            dtype=str
        )

    return pd.read_excel(
        uploaded_file,
        dtype=str
    )


def normalize_columns(df):
    """
    Normalisasi nama kolom.
    """

    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return df


def clean_text(value):
    if pd.isna(value):
        return ""

    return str(value).strip()


def parse_nominal(value):
    """
    Parser nominal yang tahan terhadap:
    1000000
    1.000.000
    1,000,000
    Rp 1.000.000
    """

    if pd.isna(value):
        return 0.0

    s = str(value).strip()

    if not s:
        return 0.0

    # Hilangkan Rp, spasi
    s = re.sub(r"(?i)rp", "", s)
    s = s.replace(" ", "")

    # Jika ada koma dan titik
    if "." in s and "," in s:

        # Indonesia: 1.000.000,50
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")

        # English: 1,000,000.50
        else:
            s = s.replace(",", "")

    # Hanya titik
    elif "." in s:

        # Jika titik terlihat sebagai separator ribuan
        parts = s.split(".")

        if all(len(x) == 3 for x in parts[1:]):
            s = "".join(parts)

    # Hanya koma
    elif "," in s:

        parts = s.split(",")

        if all(len(x) == 3 for x in parts[1:]):
            s = "".join(parts)
        else:
            s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return 0.0


def extract_va(text):
    """
    Mengenali:
    57888xxxxxxxx
    57708xxxxxxxx

    Prefix:
    57888 = BRIVA FASTPAY
    57708 = BRIVA RAJABILLER
    """

    if pd.isna(text):
        return None

    text = str(text)

    # Hilangkan karakter aneh yang mungkin mengganggu
    text = text.strip()

    match = re.search(
        r"(57888\d{5,15}|57708\d{5,15})",
        text
    )

    if match:
        return match.group(1)

    return None


def classify_va(va):

    if not va:
        return "INVALID"

    va = str(va)

    if va.startswith("57888"):
        return "BRIVA FASTPAY"

    if va.startswith("57708"):
        return "BRIVA RAJABILLER"

    return "INVALID"


def find_column(df, candidates):

    for col in candidates:
        if col in df.columns:
            return col

    return None


# ============================================================
# 1. PENGATURAN BANK
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


# Reset ketika bank berubah

if pilihan_bank != st.session_state.pilihan_bank_terakhir:

    st.session_state.sudah_diproses = False

    st.session_state.df_matched = pd.DataFrame()
    st.session_state.df_selisih_int = pd.DataFrame()
    st.session_state.df_selisih_bnk = pd.DataFrame()
    st.session_state.df_invalid_int = pd.DataFrame()
    st.session_state.df_invalid_bnk = pd.DataFrame()

    st.session_state.pilihan_bank_terakhir = pilihan_bank


# ============================================================
# 2. UPLOAD
# ============================================================

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
            key="briva_fastpay"
        )

    with col3:

        file_bnk_rajabiller = st.file_uploader(
            "🏦 BRIVA Rajabiller — 57708",
            type=["csv", "xlsx"],
            key="briva_rajabiller"
        )

else:

    col1, col2 = st.columns(2)

    with col1:

        file_int = st.file_uploader(
            "📄 FMSS",
            type=["csv", "xlsx"],
            key="fmss_general"
        )

    with col2:

        file_bnk = st.file_uploader(
            f"🏦 Mutasi Bank {pilihan_bank}",
            type=["csv", "xlsx"],
            key="bank_general"
        )


# ============================================================
# 3. BRIVA
# ============================================================

if (
    pilihan_bank == "BRIVA"
    and file_int
    and file_bnk_fastpay
    and file_bnk_rajabiller
):

    st.divider()

    st.subheader("3. Konfigurasi Fee")

    c1, c2 = st.columns(2)

    with c1:

        fee_fastpay = st.number_input(
            "Fee Fastpay (57888)",
            min_value=0,
            value=1000,
            step=100
        )

    with c2:

        fee_rajabiller = st.number_input(
            "Fee Rajabiller (57708)",
            min_value=0,
            value=1000,
            step=100
        )

    st.caption(
        "Rumus matching: Nominal FMSS + Fee = Nominal Kredit Bank"
    )

    if st.button(
        "🚀 Mulai Croscek Data BRIVA",
        type="primary"
    ):

        try:

            with st.spinner(
                "Sedang memproses rekonsiliasi BRIVA..."
            ):

                # ====================================================
                # READ FILE
                # ====================================================

                df_int = normalize_columns(
                    read_file(file_int)
                )

                df_fastpay = normalize_columns(
                    read_file(file_bnk_fastpay)
                )

                df_rajabiller = normalize_columns(
                    read_file(file_bnk_rajabiller)
                )


                # ====================================================
                # VALIDASI KOLOM FMSS
                # ====================================================

                required_int = [
                    "STATUS",
                    "KETERANGAN",
                    "NOMINAL"
                ]

                missing_int = [
                    x for x in required_int
                    if x not in df_int.columns
                ]

                if missing_int:

                    raise Exception(
                        "Kolom FMSS tidak ditemukan: "
                        + ", ".join(missing_int)
                    )


                # ====================================================
                # VALIDASI KOLOM BANK
                # ====================================================

                def prepare_bank(df, source):

                    desk_col = find_column(
                        df,
                        [
                            "DESK_TRAN",
                            "DESCRIPTION",
                            "KETERANGAN",
                            "DESKRIPSI",
                            "REMARK",
                            "NARRATIVE"
                        ]
                    )

                    kredit_col = find_column(
                        df,
                        [
                            "MUTASI_KREDIT",
                            "KREDIT",
                            "CREDIT",
                            "MUTASI CREDIT",
                            "AMOUNT"
                        ]
                    )

                    if desk_col is None:

                        raise Exception(
                            f"Kolom deskripsi pada file "
                            f"{source} tidak ditemukan."
                        )

                    if kredit_col is None:

                        raise Exception(
                            f"Kolom kredit pada file "
                            f"{source} tidak ditemukan."
                        )

                    result = df.copy()

                    result["BANK_SOURCE"] = source

                    result["BANK_DESK"] = (
                        result[desk_col]
                        .astype(str)
                        .str.strip()
                    )

                    result["BANK_NOMINAL"] = (
                        result[kredit_col]
                        .apply(parse_nominal)
                    )

                    result["VA"] = (
                        result["BANK_DESK"]
                        .apply(extract_va)
                    )

                    result["JENIS_VA"] = (
                        result["VA"]
                        .apply(classify_va)
                    )

                    return result


                # ====================================================
                # PREPARE BANK
                # ====================================================

                bank_fastpay = prepare_bank(
                    df_fastpay,
                    "BRIVA FASTPAY"
                )

                bank_rajabiller = prepare_bank(
                    df_rajabiller,
                    "BRIVA RAJABILLER"
                )


                # ====================================================
                # COMBINE BANK
                # ====================================================

                df_bank = pd.concat(
                    [
                        bank_fastpay,
                        bank_rajabiller
                    ],
                    ignore_index=True
                )


                # ====================================================
                # FMSS SUKSES
                # ====================================================

                df_int = df_int[
                    df_int["STATUS"]
                    .astype(str)
                    .str.upper()
                    .str.strip()
                    == "SUKSES"
                ].copy()


                # ====================================================
                # EXTRACT VA FMSS
                # ====================================================

                df_int["VA"] = (
                    df_int["KETERANGAN"]
                    .apply(extract_va)
                )

                df_int["JENIS_VA"] = (
                    df_int["VA"]
                    .apply(classify_va)
                )


                # ====================================================
                # NOMINAL FMSS
                # ====================================================

                df_int["NOMINAL_FMSS"] = (
                    df_int["NOMINAL"]
                    .apply(parse_nominal)
                )


                # ====================================================
                # FEE BERDASARKAN JENIS VA
                # ====================================================

                def calculate_fee(jenis):

                    if jenis == "BRIVA FASTPAY":
                        return float(fee_fastpay)

                    if jenis == "BRIVA RAJABILLER":
                        return float(fee_rajabiller)

                    return 0.0


                df_int["FEE"] = (
                    df_int["JENIS_VA"]
                    .apply(calculate_fee)
                )

                df_int["NOMINAL_MATCH"] = (
                    df_int["NOMINAL_FMSS"]
                    + df_int["FEE"]
                )


                # ====================================================
                # INVALID VA
                # ====================================================

                invalid_int = df_int[
                    df_int["VA"].isna()
                ].copy()

                invalid_bank = df_bank[
                    df_bank["VA"].isna()
                ].copy()


                # Hanya data yang memiliki VA valid
                int_valid = df_int[
                    df_int["VA"].notna()
                ].copy()

                bank_valid = df_bank[
                    df_bank["VA"].notna()
                    &
                    (df_bank["BANK_NOMINAL"] > 0)
                ].copy()


                # ====================================================
                # 1-TO-1 MATCHING
                # ====================================================

                # Index bank berdasarkan:
                # VA + NOMINAL

                bank_pool = defaultdict(list)

                for idx, row in bank_valid.iterrows():

                    key = (
                        str(row["VA"]),
                        round(
                            float(row["BANK_NOMINAL"]),
                            2
                        )
                    )

                    bank_pool[key].append(idx)


                matched = []
                unmatched_int = []

                used_bank_index = set()


                for _, int_row in int_valid.iterrows():

                    key = (
                        str(int_row["VA"]),
                        round(
                            float(int_row["NOMINAL_MATCH"]),
                            2
                        )
                    )

                    candidates = bank_pool.get(
                        key,
                        []
                    )

                    # Cari bank row yang belum dipakai

                    selected_idx = None

                    for candidate in candidates:

                        if candidate not in used_bank_index:

                            selected_idx = candidate
                            break


                    # ==================================================
                    # MATCH
                    # ==================================================

                    if selected_idx is not None:

                        bank_row = bank_valid.loc[
                            selected_idx
                        ]

                        used_bank_index.add(
                            selected_idx
                        )

                        record = int_row.to_dict()

                        record["MATCH_BANK_SOURCE"] = (
                            bank_row["BANK_SOURCE"]
                        )

                        record["MATCH_BANK_VA"] = (
                            bank_row["VA"]
                        )

                        record["MATCH_BANK_NOMINAL"] = (
                            bank_row["BANK_NOMINAL"]
                        )

                        record["MATCH_BANK_DESK"] = (
                            bank_row["BANK_DESK"]
                        )

                        record["MATCH_STATUS"] = "MATCHED"

                        matched.append(record)

                    else:

                        record = int_row.to_dict()

                        record["MATCH_STATUS"] = (
                            "FMSS_ONLY"
                        )

                        unmatched_int.append(record)


                # ====================================================
                # BANK UNMATCHED
                # ====================================================

                unmatched_bank = bank_valid[
                    ~bank_valid.index.isin(
                        used_bank_index
                    )
                ].copy()


                unmatched_bank["ISSUE"] = "BANK_ONLY"


                # ====================================================
                # OUTPUT
                # ====================================================

                df_matched = pd.DataFrame(
                    matched
                )

                df_selisih_int = pd.DataFrame(
                    unmatched_int
                )

                df_selisih_bnk = (
                    unmatched_bank
                    .copy()
                )


                # ====================================================
                # TAMBAHKAN ISSUE
                # ====================================================

                if not df_selisih_int.empty:

                    df_selisih_int["ISSUE"] = (
                        "FMSS_ONLY"
                    )


                # ====================================================
                # SAVE SESSION
                # ====================================================

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
                    invalid_int
                )

                st.session_state.df_invalid_bnk = (
                    invalid_bank
                )

                st.session_state.sudah_diproses = True


        except Exception as e:

            st.error(
                f"❌ Gagal memproses data: {str(e)}"
            )

            st.session_state.sudah_diproses = False


# ============================================================
# 4. HASIL BRIVA
# ============================================================

if (
    pilihan_bank == "BRIVA"
    and st.session_state.sudah_diproses
):

    df_matched = st.session_state.df_matched
    df_selisih_int = st.session_state.df_selisih_int
    df_selisih_bnk = st.session_state.df_selisih_bnk
    df_invalid_int = st.session_state.df_invalid_int
    df_invalid_bnk = st.session_state.df_invalid_bnk


    # ========================================================
    # SUMMARY TRANSAKSI
    # ========================================================

    st.divider()

    st.subheader(
        "🎯 Ringkasan Rekonsiliasi BRIVA"
    )

    m1, m2, m3, m4 = st.columns(4)

    m1.metric(
        "✅ Matched Sempurna",
        f"{len(df_matched):,} Trx"
    )

    m2.metric(
        "⚠️ Issue FMSS",
        f"{len(df_selisih_int):,} Trx"
    )

    m3.metric(
        "⚠️ Issue Bank",
        f"{len(df_selisih_bnk):,} Trx"
    )

    m4.metric(
        "🚨 Invalid VA",
        f"{len(df_invalid_int) + len(df_invalid_bnk):,} Trx"
    )


    # ========================================================
    # SUMMARY NOMINAL
    # ========================================================

    st.subheader(
        "💰 Ringkasan Nominal"
    )

    matched_nominal = (
        df_matched["NOMINAL_FMSS"].sum()
        if not df_matched.empty
        else 0
    )

    issue_int_nominal = (
        df_selisih_int["NOMINAL_FMSS"].sum()
        if not df_selisih_int.empty
        else 0
    )

    issue_bank_nominal = (
        df_selisih_bnk["BANK_NOMINAL"].sum()
        if not df_selisih_bnk.empty
        else 0
    )

    n1, n2, n3 = st.columns(3)

    n1.metric(
        "Matched",
        f"Rp {matched_nominal:,.0f}"
    )

    n2.metric(
        "Issue FMSS",
        f"Rp {issue_int_nominal:,.0f}"
    )

    n3.metric(
        "Issue Bank",
        f"Rp {issue_bank_nominal:,.0f}"
    )


    # ========================================================
    # BREAKDOWN JENIS VA
    # ========================================================

    st.divider()

    st.subheader(
        "🏦 Breakdown BRIVA"
    )

    b1, b2 = st.columns(2)

    with b1:

        st.markdown(
            "#### BRIVA Fastpay — 57888"
        )

        fast_matched = (
            df_matched[
                df_matched["JENIS_VA"]
                == "BRIVA FASTPAY"
            ]
            if not df_matched.empty
            else pd.DataFrame()
        )

        st.metric(
            "Matched",
            f"{len(fast_matched):,} Trx"
        )

    with b2:

        st.markdown(
            "#### BRIVA Rajabiller — 57708"
        )

        raj_matched = (
            df_matched[
                df_matched["JENIS_VA"]
                == "BRIVA RAJABILLER"
            ]
            if not df_matched.empty
            else pd.DataFrame()
        )

        st.metric(
            "Matched",
            f"{len(raj_matched):,} Trx"
        )


    # ========================================================
    # ISSUE TABLE
    # ========================================================

    st.divider()

    c1, c2 = st.columns(2)

    with c1:

        st.markdown(
            "### 🚨 Issue FMSS"
        )

        if not df_selisih_int.empty:

            tampil = df_selisih_int[
                [
                    "VA",
                    "JENIS_VA",
                    "NOMINAL_FMSS",
                    "FEE",
                    "NOMINAL_MATCH",
                    "ISSUE"
                ]
            ].copy()

            tampil.columns = [
                "KODE VA",
                "JENIS VA",
                "NOMINAL",
                "FEE",
                "NOMINAL MATCH",
                "ISSUE"
            ]

            st.dataframe(
                tampil,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.success(
                "Tidak ada issue FMSS."
            )


    with c2:

        st.markdown(
            "### 🚨 Issue Bank"
        )

        if not df_selisih_bnk.empty:

            tampil = df_selisih_bnk[
                [
                    "VA",
                    "JENIS_VA",
                    "BANK_NOMINAL",
                    "BANK_SOURCE",
                    "ISSUE"
                ]
            ].copy()

            tampil.columns = [
                "KODE VA",
                "JENIS VA",
                "NOMINAL",
                "SUMBER BANK",
                "ISSUE"
            ]

            st.dataframe(
                tampil,
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

    if (
        not df_invalid_int.empty
        or not df_invalid_bnk.empty
    ):

        with st.expander(
            "⚠️ Transaksi dengan VA Tidak Teridentifikasi"
        ):

            c1, c2 = st.columns(2)

            with c1:

                st.markdown(
                    "#### FMSS Invalid VA"
                )

                if not df_invalid_int.empty:

                    st.dataframe(
                        df_invalid_int[
                            [
                                "KETERANGAN",
                                "NOMINAL_FMSS"
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True
                    )

                else:

                    st.success(
                        "Tidak ada FMSS invalid VA."
                    )

            with c2:

                st.markdown(
                    "#### Bank Invalid VA"
                )

                if not df_invalid_bnk.empty:

                    st.dataframe(
                        df_invalid_bnk[
                            [
                                "BANK_SOURCE",
                                "BANK_DESK",
                                "BANK_NOMINAL"
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True
                    )

                else:

                    st.success(
                        "Tidak ada Bank invalid VA."
                    )


    # ========================================================
    # DOWNLOAD EXCEL
    # ========================================================

    st.divider()

    st.subheader(
        "📥 Download Laporan"
    )

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        if not df_matched.empty:

            df_matched.to_excel(
                writer,
                sheet_name="MATCHED_OK",
                index=False
            )

        else:

            pd.DataFrame(
                {
                    "INFO": [
                        "Tidak ada data matched."
                    ]
                }
            ).to_excel(
                writer,
                sheet_name="MATCHED_OK",
                index=False
            )


        if not df_selisih_int.empty:

            df_selisih_int.to_excel(
                writer,
                sheet_name="ISSUE_FMSS",
                index=False
            )

        else:

            pd.DataFrame(
                {
                    "INFO": [
                        "Tidak ada issue FMSS."
                    ]
                }
            ).to_excel(
                writer,
                sheet_name="ISSUE_FMSS",
                index=False
            )


        if not df_selisih_bnk.empty:

            df_selisih_bnk.to_excel(
                writer,
                sheet_name="ISSUE_BANK",
                index=False
            )

        else:

            pd.DataFrame(
                {
                    "INFO": [
                        "Tidak ada issue Bank."
                    ]
                }
            ).to_excel(
                writer,
                sheet_name="ISSUE_BANK",
                index=False
            )


        if not df_invalid_int.empty:

            df_invalid_int.to_excel(
                writer,
                sheet_name="INVALID_FMSS",
                index=False
            )


        if not df_invalid_bnk.empty:

            df_invalid_bnk.to_excel(
                writer,
                sheet_name="INVALID_BANK",
                index=False
            )


    st.download_button(
        label="📥 Download Laporan Lengkap (.xlsx)",
        data=output.getvalue(),
        file_name="Laporan_Rekonsiliasi_BRIVA.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary"
    )


# ============================================================
# 5. BANK LAIN
# ============================================================

elif (
    pilihan_bank != ""
    and pilihan_bank != "BRIVA"
    and file_int
    and file_bnk
):

    st.divider()

    st.warning(
        f"🚧 Modul rekonsiliasi {pilihan_bank} "
        "belum diaktifkan pada versi ini."
    )


elif pilihan_bank == "":

    st.info(
        "💡 Silakan pilih Bank Sumber Mutasi terlebih dahulu."
    )
