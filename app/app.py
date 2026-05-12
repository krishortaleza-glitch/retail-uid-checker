import streamlit as st
import pandas as pd
import tempfile

st.set_page_config(page_title="Retail UID Checker", layout="wide")

st.title("🔍 Retail UID Difference Checker")

# =========================
# LOAD FILES
# =========================
master_file = st.file_uploader(
    "Upload Master Product File",
    type=["xlsx", "csv"]
)

bu_file = st.file_uploader(
    "Upload BU Product File",
    type=["xlsx", "csv"]
)

# =========================
# HELPERS
# =========================
def load_file(file):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file, dtype=str)
    else:
        df = pd.read_excel(file, dtype=str)

    df.columns = df.columns.str.strip()
    return df

def clean_upc(series):
    return (
        series.astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\D", "", regex=True)
        .str.strip()
    )

# =========================
# PROCESS
# =========================
if master_file and bu_file:

    master_df = load_file(master_file)
    bu_df = load_file(bu_file)

    st.success("Files Loaded")

    st.subheader("Map Columns")

    # MASTER
    master_upc_col = st.selectbox(
        "Master UPC Column",
        master_df.columns,
        key="master_upc"
    )

    master_uid_col = st.selectbox(
        "Master Retail UID Column",
        master_df.columns,
        key="master_uid"
    )

    # BU
    bu_upc_col = st.selectbox(
        "BU UPC Column",
        bu_df.columns,
        key="bu_upc"
    )

    bu_uid_col = st.selectbox(
        "BU Retail UID Column",
        bu_df.columns,
        key="bu_uid"
    )

    if st.button("🚀 Compare Files"):

        # CLEAN UPCS
        master_df["UPC_CLEAN"] = clean_upc(master_df[master_upc_col])
        bu_df["UPC_CLEAN"] = clean_upc(bu_df[bu_upc_col])

        # BUILD MASTER LOOKUP
        master_lookup = (
            master_df[
                ["UPC_CLEAN", master_uid_col]
            ]
            .drop_duplicates()
            .rename(columns={
                master_uid_col: "MASTER_UID"
            })
        )

        # MERGE
        merged = bu_df.merge(
            master_lookup,
            on="UPC_CLEAN",
            how="left"
        )

        # FLAG DIFFERENCES
        flagged = merged[
            (merged["MASTER_UID"].notna()) &
            (
                merged[bu_uid_col].astype(str).str.strip()
                !=
                merged["MASTER_UID"].astype(str).str.strip()
            )
        ].copy()

        flagged["BU_UID"] = flagged[bu_uid_col]

        # OUTPUT
        output_cols = [
            bu_upc_col,
            "BU_UID",
            "MASTER_UID"
        ]

        remaining_cols = [
            c for c in flagged.columns
            if c not in output_cols
            and c != "UPC_CLEAN"
        ]

        flagged = flagged[
            output_cols + remaining_cols
        ]

        st.success(
            f"Found {len(flagged)} Retail UID mismatches"
        )

        st.dataframe(flagged)

        # EXPORT
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as tmp:

            temp_path = tmp.name

        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            flagged.to_excel(
                writer,
                sheet_name="UID Mismatches",
                index=False
            )

        with open(temp_path, "rb") as f:
            file_bytes = f.read()

        st.download_button(
            "⬇ Download Flagged Items",
            data=file_bytes,
            file_name="retail_uid_mismatches.xlsx"
        )
