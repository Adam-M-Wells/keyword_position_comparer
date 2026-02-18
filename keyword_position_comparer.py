import streamlit as st
import pandas as pd
from functools import reduce
from io import BytesIO

# -------------------------
# Page setup
# -------------------------
st.set_page_config(page_title="Keyword Position Comparer", layout="wide")
st.title("📊 Keyword Position Comparer")

st.markdown("""
Use this tool to compare your client's rankings to competitors or compare ranking snapshots.

### How to use:

📂 **Upload 3 to 6 .xls or .xlsx spreadsheets** with columns in this order:

1. Keyword  
2. Position  
3. Search Volume  
4. CPC  
5. URL  

⚠️ **First file = client spreadsheet**
""")

# -------------------------
# Upload
# -------------------------
uploaded_files = st.file_uploader(
    "Upload 3 to 6 spreadsheets",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    if len(uploaded_files) < 3:
        st.warning("Please upload at least 3 spreadsheets.")
        st.stop()

    if len(uploaded_files) > 6:
        st.warning("Please upload no more than 6 spreadsheets.")
        st.stop()

    st.info(f"Client file detected as: {uploaded_files[0].name}")

    dfs = []
    num_files = len(uploaded_files)

    # -------------------------
    # Read files safely
    # -------------------------
    for i, file in enumerate(uploaded_files, start=1):
        try:
            df = pd.read_excel(file)

            # Validate structure
            if df.shape[1] < 5:
                st.error(f"{file.name} has fewer than 5 columns.")
                continue

            df = df.iloc[:, :5]
            df.columns = [
                "Keyword",
                f"Position_{i}",
                f"Search Volume_{i}",
                f"CPC_{i}",
                f"URL_{i}"
            ]

            # FORCE Keyword column to string and strip spaces
            df["Keyword"] = df["Keyword"].astype(str).str.strip()

            # Remove empty keywords
            df = df[df["Keyword"] != ""]

            dfs.append(df)

        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")

    # Stop if nothing loaded
    if len(dfs) == 0:
        st.error("No files could be processed. Check file format.")
        st.stop()

    # -------------------------
    # Keyword reference set
    # -------------------------
    keywords_in_first = set(dfs[0]["Keyword"].dropna())

    # -------------------------
    # Merge
    # -------------------------
    merged_df = reduce(
        lambda left, right: pd.merge(left, right, on="Keyword", how="outer"),
        dfs
    )

    # -------------------------
    # Fill missing values
    # -------------------------
    for i in range(1, num_files + 1):
        merged_df[f"Position_{i}"] = merged_df[f"Position_{i}"].fillna("N/A")
        merged_df[f"URL_{i}"] = merged_df[f"URL_{i}"].fillna("N/A")

    # -------------------------
    # First non-null Search Volume & CPC
    # -------------------------
    merged_df["Search Volume"] = (
        merged_df[[f"Search Volume_{i}" for i in range(1, num_files + 1)]]
        .bfill(axis=1)
        .iloc[:, 0]
        .fillna("N/A")
    )

    merged_df["CPC"] = (
        merged_df[[f"CPC_{i}" for i in range(1, num_files + 1)]]
        .bfill(axis=1)
        .iloc[:, 0]
        .fillna("N/A")
    )

    # -------------------------
    # Appearance count
    # -------------------------
    position_cols = [f"Position_{i}" for i in range(1, num_files + 1)]
    merged_df["Appearances"] = merged_df[position_cols].apply(
        lambda row: sum(x != "N/A" for x in row),
        axis=1
    )

    # -------------------------
    # Build output dataframe
    # -------------------------
    output_df = pd.DataFrame()
    output_df["Keyword"] = merged_df["Keyword"]

    for i in range(1, num_files + 1):
        output_df[f"Position from Spreadsheet {i}"] = merged_df[f"Position_{i}"]

    output_df["Search Volume"] = merged_df["Search Volume"]
    output_df["CPC"] = merged_df["CPC"]

    for i in range(1, num_files + 1):
        output_df[f"URL from Spreadsheet {i}"] = merged_df[f"URL_{i}"]

    output_df["Appearances"] = merged_df["Appearances"]

    # -------------------------
    # Split tabs
    # -------------------------
    tab1 = output_df[
        output_df["Keyword"].isin(keywords_in_first)
    ].drop(columns=["Appearances"])

    tab2 = output_df[
        (~output_df["Keyword"].isin(keywords_in_first))
        & (output_df["Appearances"] >= 2)
    ].drop(columns=["Appearances"])

    tab3 = output_df[
        (~output_df["Keyword"].isin(keywords_in_first))
        & (output_df["Appearances"] == 1)
    ].drop(columns=["Appearances"])

    # -------------------------
    # Excel export
    # -------------------------
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        tab1.to_excel(writer, sheet_name="Client", index=False)
        tab2.to_excel(writer, sheet_name="2+ Competitors", index=False)
        tab3.to_excel(writer, sheet_name="1 Competitor", index=False)

    output.seek(0)

    # -------------------------
    # Download button
    # -------------------------
    st.success("Files processed successfully!")

    st.download_button(
        label="📥 Download Combined Spreadsheet",
        data=output,
        file_name="combined_keywords_split.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # -------------------------
    # Preview tables
    # -------------------------
    st.subheader("Preview: Client Keywords")
    st.dataframe(tab1.head(20))

    st.subheader("Preview: 2+ Competitors (Not in Client)")
    st.dataframe(tab2.head(20))

    st.subheader("Preview: Only 1 Competitor (Not in Client)")
    st.dataframe(tab3.head(20))
