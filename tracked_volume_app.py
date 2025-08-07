import streamlit as st
import pandas as pd
import io

st.title("Tracked Volume Filter App")

st.write(
    "Upload your data Excel file. This app will filter columns for 'Tracked Volume' and let you download the result."
)

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    tracked_cols = [
        i
        for i, val in enumerate(df.iloc[1])
        if str(val).strip().lower() == "tracked volume"
    ]
    cols_to_keep = [0, 1, 2, 3] + tracked_cols

    # Prepare rows for output
    month_row = df.iloc[[0], cols_to_keep]
    header_row = df.iloc[[1], cols_to_keep]
    data_rows = df.iloc[2:, cols_to_keep]

    final_output = pd.concat([month_row, header_row, data_rows], ignore_index=True)

    # Convert DataFrame to Excel in-memory
    output = io.BytesIO()
    final_output.to_excel(output, index=False, header=False)
    output.seek(0)

    st.success("Filtering done! Click below to download the result.")
    st.download_button(
        label="Download Filtered Excel",
        data=output,
        file_name="Filtered_Tracked_Volume.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.write("Preview of output:")
    st.dataframe(final_output.head(10))  # Show first 10 rows as preview
