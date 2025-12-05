import streamlit as st
import pandas as pd
import numpy as np
import time
import re
from io import BytesIO

def format_rupiah(x):
    if pd.isna(x):
        return ""
    # pastikan bisa diubah ke float
    try:
        x = float(x)
    except:
        return x  # biarin apa adanya kalau bukan angka

    # kalau tidak punya desimal (misal 7000.0), tampilkan tanpa ,00
    if x.is_integer():
        formatted = f"{int(x):,}".replace(",", ".")
    else:
        formatted = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # hapus ,00 kalau desimalnya 0 semua (misal 7000,00 → 7000)
        if formatted.endswith(",00"):
            formatted = formatted[:-3]
    return formatted

def highlight_total(row):
    # Cek apakah ada kolom yang berisi "TOTAL" (case-insensitive)
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return ["font-weight: bold; background-color: #D9EAD3; color: #1A5E20;"] * len(row)
    else:
        return [""] * len(row)
    
def highlight_1st_2nd_vendor(row, columns):
    styles = [""] * len(columns)
    first_vendor = row.get("1st Vendor")
    second_vendor = row.get("2nd Vendor")

    for i, col in enumerate(columns):
        if col == first_vendor:
            # styles[i] = "background-color: #f8c8dc; color: #7a1f47;"
            styles[i] = "background-color: #C6EFCE; color: #006100;"
        elif col == second_vendor:
            # styles[i] = "background-color: #d7c6f3; color: #402e72;"
            styles[i] = "background-color: #FFEB9C; color: #9C6500;"
    return styles

st.subheader("🧑‍🏫 User Guide: TCO Comparison by Region")
st.markdown(
    ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
)
st.caption("INSPIRE 2025 | Oktaviana Sadama Nur Azizah")

# Divider custom
st.markdown(
    """
    <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="
        display: flex;
        align-items: center;
        height: 65px;
        margin-bottom: 10px;
    ">
        <div style="text-align: justify; font-size: 15px;">
            <span style="color: #FF2EC4; font-weight: 800;">
            TCO Comparison by Region</span>
            analyzes TCO differences across regions to highlight cost variations based on geographic 
            requirements.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("#### Input Structure")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The input file required for this menu should be a 
            <span style="color: #FF69B4; font-weight: 500;">single file containing multiple sheets</span>, in eather 
            <span style="background:#C6EFCE; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xlsx</span> or 
            <span style="background:#FFEB9C; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xls</span> format. 
            Each sheet represents a vendor name, with the table structure in each sheet as follows:
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["Scope", "Desc", "Region 1", "Region 2", "Region 3", "Region 4", "Region 5"]
df = pd.DataFrame([[""] * len(columns) for _ in range(3)], columns=columns)

st.dataframe(df, hide_index=True)

# Buat DataFrame 1 row
st.markdown("""
<table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 15px;">
    <tr>
        <td style="border: 1px solid gray; width: 15%;">Vendor A</td>
        <td style="border: 1px solid gray; width: 15%;">Vendor B</td>
        <td style="border: 1px solid gray; width: 15%;">Vendor C</td>
        <td style="border: 1px solid gray; font-style: italic; color: #26BDAD">multiple sheets</td>
    </tr>
</table>
""", unsafe_allow_html=True)

st.markdown("###### Description:")
st.markdown(
    """
    <div style="font-size:15px;">
        <ul>
            <li>
                <span style="display:inline-block; width:100px;">Scope & Desc</span>: non-numeric columns
            </li>
            <li>
                <span style="display:inline-block; width:100px;">Region 1 to 5</span>: numeric columns
            </li>
        </ul>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The system accommodates a 
            <span style="font-weight: bold;">dynamic table</span>, 
            allowing users to enter any number of non-numeric and numeric columns. 
            Users have the freedom to name the columns as they wish. The system logic relies on 
            <span style="font-weight: bold;">column indices</span>, not specific column names.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:violet-badge[Ensure that each sheet has the same table structure and column names!]**")

st.divider()
st.markdown("#### Constraint")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -10px">
            To ensure this menu works correctly, users need to follow certain rules regarding
            the dataset structure.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. COLUMN ORDER]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top: -10px">
            When creating tables, it is important to follow the specified column structure. Columns 
            <span style="font-weight: bold;">must</span> be arranged in the following order:
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: center; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            Non-Numeric Columns → Numeric Columns
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px">
            this order is <span style="color: #FF69B4; font-weight: 700;">strict</span> and 
            <span style="color: #FF69B4; font-weight: 700;">cannot be altered</span>!
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:orange-badge[2. NUMBER COLUMN]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Please refer the table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["No", "Scope", "Desc", "Region 1", "Region 2", "Region 3", "Region 4", "Region 5"]
data = [
    [1] + [""] * (len(columns) - 1),
    [2] + [""] * (len(columns) - 1),
    [3] + [""] * (len(columns) - 1)
]
df = pd.DataFrame(data, columns=columns)

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top: -5px;">
            The table above is an 
            <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
            <span style="color: #FF69B4; font-weight: 700;">not allowed</span> because it contains a 
            <span style="font-weight: bold;">"No"</span> column. 
            The "No" column is prohibited in this menu, as it will be treated as a numeric column by the system, 
            which violates the constraint described in point 1.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:green-badge[3. FLOATING TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Floating tables are allowed, meaning tables 
            <span style="color: #FF69B4; font-weight: 700;">do not need to start from cell A1</span>. 
            However, ensure that the cells above and to the left of the table are empty, as shown in the example below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["", "A", "B", "C", "D", "E", "F"]

# Buat 5 baris kosong
df = pd.DataFrame([[""] * len(columns) for _ in range(6)], columns=columns)

# Isi kolom pertama dengan 1–6
df.iloc[:, 0] = [1, 2, 3, 4, 5, 6]

# Header bagian kedua
df.loc[1, ["B", "C", "D", "E"]] = ["Scope", "Region 1", "Region 2", "Region 3"]

# Data Software & Hardware
df.loc[2, ["B", "C", "D", "E"]] = ["MBTS", "1.000", "2.000", "3.000"]
df.loc[3, ["B", "C", "D", "E"]] = ["Reposition", "5.500", "6.500", "7.500"]
df.loc[4, ["B", "C", "D", "E"]] = ["Reroute", "1.200", "1.700", "2.200"]

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top:-10px;">
            To provide additional explanations or notes on the sheet, you can include them using an image or a text box.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:blue-badge[4. TOTAL COLUMN & TOTAL ROW]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            You are not allowed to add a
            <span style="font-weight: 700;">TOTAL COLUMN</span> or
            <span style="font-weight: 700;">TOTAL ROW</span>!
            Please refer to the example table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Scope", "Region 1", "Region 2", "Region 3", "TOTAL"]
data = [
    ["MBTS", "1.000", "2.000", "3.000", "6.000"],
    ["Reroute", "1.200", "1.700", "2.200", "5.100"],
    ["TOTAL", "2.200", "3.700", "5.200", "11.100"],
]
df = pd.DataFrame(data, columns=columns)

def red_highlight(row):
    styles = [""] * len(row)

    # Highlight ROW "TOTAL"
    if row["Scope"] == "TOTAL":
        styles = ["color: #FF4D4D;" for _ in row]
    else:
        # Highlight COLUMN "TOTAL"
        total_col_index = row.index.get_loc("TOTAL")
        styles[total_col_index] = "color: #FF4D4D;"

    return styles

num_cols = ["Y0", "Y1", "Y2", "Y3", "TOTAL 3Y TCO"]
df_styled = df.style.apply(red_highlight, axis=1)

st.dataframe(df_styled, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -5px;">
            The table above is an 
            <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
            <span style="color: #FF69B4; font-weight: 700;">not permitted</span>! 
            The total column & row are generated automatically during
            <span style="font-weight: 700;">MERGE DATA</span> — 
            do not add them manually. If added, the system will treat them as part of the region & scope, and included them in calculations.
        </div>
    """,
    unsafe_allow_html=True
)

st.divider()

st.markdown("#### Note")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top:-10px;">
            This menu has two main focuses: analyzing the 
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">ORIGINAL DATA</span> and the 
            <span style="background: #FF00AA; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">TRANSPOSED DATA</span>.
            Each displayed in separate tabs, as shown below.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["ORIGINAL DATA", "TRANSPOSE DATA"])

st.markdown("**:gray-badge[Those tabs appears immediately after the user uploads the input!]**")

st.divider()

st.markdown("#### What is Displayed?")

# Path file Excel yang sudah ada
file_path = "dummy dataset.xlsx"

# Buka file sebagai binary
with open(file_path, "rb") as f:
    file_data = f.read()

# Markdown teks
st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 5px; margin-top: -10px">
        You can try this menu by downloading the dummy dataset using the button below: 
    </div>
    """,
    unsafe_allow_html=True
)

@st.fragment
def release_the_balloons():
    st.balloons()

# Download button untuk file Excel
st.download_button(
    label="Dummy Dataset",
    data=file_data,
    file_name="dummy dataset.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    on_click=release_the_balloons,
    type="primary",
    use_container_width=True,
)

st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
        Based on this dummy dataset, the menu will produce the following results.
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. MERGE DATA]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system will merge the tables from each sheet into a single table and add a 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">TOTAL ROW</span> 
            for each vendor.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            🛸 Original Data
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["VENDOR", "SCOPE TOTAL PRICE (IDR)", "REGION 1", "REGION 2", "REGION 3", "TOTAL"]
data = [
    ["Vendor A", "MBTS", 800, 250, 300, 1350],
    ["Vendor A", "Reposition", 750, 250, 260, 1260],
    ["Vendor A", "Reroute", 1140, 380, 390, 1910],
    ["Vendor A", "TOTAL", 2690, 880, 950, 4520],

    ["Vendor B", "MBTS", 1250, 400, 450, 2100],
    ["Vendor B", "Reposition", 650, 200, 220, 1070],
    ["Vendor B", "Reroute", 810, 270, 280, 1360],
    ["Vendor B", "TOTAL", 2710, 870, 950, 4530],

    ["Vendor C", "MBTS", 900, 300, 320, 1520],
    ["Vendor C", "Reposition", 980, 320, 350, 1650],
    ["Vendor C", "Reroute", 720, 230, 240, 1190],
    ["Vendor C", "TOTAL", 2600, 850, 910, 4360],
]
df_merge = pd.DataFrame(data, columns=columns)

num_cols = ["REGION 1", "REGION 2", "TOTAL"]
df_merge_styled = (
    df_merge.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)

st.dataframe(df_merge_styled, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            👽 Transpose Data
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["VENDOR", "REGION", "MBTS", "Reposition", "Reroute", "TOTAL"]
data = [
    ["Vendor A", "REGION 1", 800,750,1140,2690],
    ["Vendor A", "REGION 2", 250,250,380,880],
    ["Vendor A", "REGION 3", 300,260,390,950],
    ["Vendor A", "TOTAL", 1350,1260,1910,4520],

    ["Vendor B", "REGION 1", 1250,650,810,2710],
    ["Vendor B", "REGION 2", 400,200,270,870],
    ["Vendor B", "REGION 3", 450,220,280,950],
    ["Vendor B", "TOTAL", 2100,1070,1360,4530],

    ["Vendor C", "REGION 1", 900,980,720,2600],
    ["Vendor C", "REGION 2", 300,320,230,850],
    ["Vendor C", "REGION 3", 320,350,240,910],
    ["Vendor C", "TOTAL", 1520,1650,1190,4360],
]
df_merge_transpose = pd.DataFrame(data, columns=columns)

num_cols = ["MBTS", "Reposition", "Reroute", "TOTAL"]
df_merge_transpose_styled = (
    df_merge_transpose.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)

st.dataframe(df_merge_transpose_styled, hide_index=True)

st.write("")
st.markdown("**:orange-badge[2. TCO SUMMARY]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            After merging the data, the system will automatically generate a TCO Summary that includes 
            the TOTAL calculations.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            🛸 Original Data
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["SCOPE TOTAL PRICE (IDR)", "VENDOR A", "VENDOR B", "VENDOR C"]
data = [
    ["MBTS", 1350, 2100, 1520],
    ["Reposition", 1260, 1070, 1650],
    ["Reroute", 1910, 1360, 1190],
    ["TOTAL", 4520, 4530, 4360]
]
df_tco_ori = pd.DataFrame(data, columns=columns)

num_cols = ["VENDOR A", "VENDOR B", "VENDOR C"]
df_tco_ori_styled = (
    df_tco_ori.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)
st.dataframe(df_tco_ori_styled, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            👽 Transposed Data
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["REGION", "VENDOR A", "VENDOR B", "VENDOR C"]
data = [
    ["REGION 1", 2690, 2710, 2600],
    ["REGION 2", 880, 870, 850],
    ["REGION 3", 950, 950, 910],
    ["TOTAL", 4520, 4530, 4360]
]
df_tco_transpose = pd.DataFrame(data, columns=columns)

num_cols = ["VENDOR A", "VENDOR B", "VENDOR C"]
df_tco_transpose_styled = (
    df_tco_transpose.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)
st.dataframe(df_tco_transpose_styled, hide_index=True)

st.write("")
st.markdown("**:yellow-badge[3. BID & PRICE ANALYSIS]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu also displays an analysis table that provides a comprehensive overview of the pricing structure 
            submitted by each vendor.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            🛸 Original Data
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="text-align:left; margin-bottom: 8px">
        <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
        &nbsp;
        <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
    </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["REGION", "SCOPE TOTAL PRICE (IDR)", "VENDOR A", "VENDOR B", "VENDOR C", "1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price", "VENDOR A to Median (%)", "VENDOR B to Median (%)", "VENDOR C to Median (%)"]
data = [
    ["REGION 1", "MBTS", 800, 1250, 900, 800, "VENDOR A", 900, "VENDOR C", "12.5%", 900, "-11.1%", "+38.9%", "+0.0%"],
    ["REGION 1", "Reposition", 750, 650, 980, 650, "VENDOR B", 750, "VENDOR A", "15.4%", 750, "+0.0%", "-13.3%", "+30.7%"],
    ["REGION 1", "Reroute", 1140, 810, 720, 720, "VENDOR C", 810, "VENDOR B", "12.5%", 810, "+40.7%", "+0.0%", "-11.1%"],

    ["REGION 2", "MBTS", 250, 400, 300, 250, 'VENDOR A', 300, "VENDOR C", "20.0%", 300, "-16.7%", "+33.3%", "+0.0%"],
    ["REGION 2", "Reposition", 250, 200, 320, 200, "VENDOR B", 250, "VENDOR A", "25.0%", 250, "+0.0%", "-20.0%" ,"+28.0%"],
    ["REGION 2", "Reroute", 380, 270, 230, 230, "VENDOR C", 270, "VENDOR B", "17.4%", 270, "+40.7%", "+0.0%", "-14.8%"],

    ["REGION 3", "MBTS", 300, 450, 320, 300, "VENDOR A", 320, "VENDOR C", "6.7%", 320, "-6.2%", "+40.6%", "+0.0%"],
    ["REGION 3", "Reposition", 260, 220, 350, 220, "VENDOR B", 260, "VENDOR A", "18.2%", 260, "+0.0%", "-15.4%", "+34.6%"],
    ["REGION 3", "Reroute", 390, 280, 240, 240, "VENDOR C", 280, "VENDOR B", "16.7%", 280, "+39.3%", "+0.0%", "-14.3%"],
]
df_analysis = pd.DataFrame(data, columns=columns)

num_cols = ["VENDOR A", "VENDOR B", "VENDOR C", "1st Lowest", "2nd Lowest", "Median Price"]
df_analysis_styled = (
    df_analysis.style
    .format({col: format_rupiah for col in num_cols})
    .apply(lambda row: highlight_1st_2nd_vendor(row, df_analysis.columns), axis=1)
)

st.dataframe(df_analysis_styled, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            👽 Transpose Data
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="text-align:left; margin-bottom: 8px">
        <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
        &nbsp;
        <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
    </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["SCOPE", "REGION", "VENDOR A", "VENDOR B", "VENDOR C", "1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price", "VENDOR A to Median (%)", "VENDOR B to Median (%)", "VENDOR C to Median (%)"]
data = [
    ["MBTS","REGION 1", 800, 1250, 900, 800, "VENDOR A", 900, "VENDOR C", "12.5%", 900, "-11.1%", "+38.9%", "+0.0%"],
    ["MBTS","REGION 2", 250, 400, 300, 250, "VENDOR A", 300, "VENDOR C", "20.0%", 300, "-16.7%", "+33.3%", "+0.0%"],
    ["MBTS","REGION 3", 300, 450, 320, 300, "VENDOR A", 320, "VENDOR C", "6.7%", 320, "-6.2%", "+40.6%", "+0.0%"],

    ["Reposition","REGION 1", 750, 650, 980, 650, 'VENDOR B', 750, "VENDOR A", "15.4%", 750, "+0.0%", "-13.3%", "+30.7%"],
    ["Reposition","REGION 2", 250, 200, 320, 200, "VENDOR B", 250, "VENDOR A", "25.0%", 250, "+0.0%", "-20.0%" ,"+28.0%"],
    ["Reposition","REGION 3", 260, 220, 350, 220, "VENDOR B", 260, "VENDOR A", "18.2%", 260, "+0.0%", "-15.4%", "+34.6%"],

    ["Reroute","REGION 1", 1140, 810, 720, 720, "VENDOR C", 810, "VENDOR B", "12.5%", 810, "+40.7%", "+0.0%", "-11.1%"],
    ["Reroute","REGION 2", 380, 270, 230, 230, "VENDOR C", 270, "VENDOR B", "17.4%", 270, "+40.7%", "+0.0%", "-14.8%"],
    ["Reroute","REGION 3", 390, 280, 240, 240, "VENDOR C", 280, "VENDOR B", "16.7%", 280, "+39.3%", "+0.0%", "-14.3%"],
]
df_analysis_transpose = pd.DataFrame(data, columns=columns)

num_cols = ["VENDOR A", "VENDOR B", "VENDOR C", "1st Lowest", "2nd Lowest", "Median Price"]
df_analysis_transpose_styled = (
    df_analysis_transpose.style
    .format({col: format_rupiah for col in num_cols})
    .apply(lambda row: highlight_1st_2nd_vendor(row, df_analysis_transpose.columns), axis=1)
)

st.dataframe(df_analysis_transpose_styled, hide_index=True)

st.write("")
st.markdown("**:green-badge[4. VISUALIZATION]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu displays visualizations focusing on two key aspects: 
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Win Rate Trend</span> and 
            <span style="background: #FF00AA; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Average Gap Trend</span>, 
            each presented in its own tab.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["Win Rate Trend", "Average Gap Trend"])

with tab1:
    st.image("assets/1.png")
    with st.expander("See explanation"):
        st.caption('''
            The visualization above compares the win rate of each vendor
            based on how often they achieved 1st or 2nd place in all
            tender evaluations.  
                    
            **💡 How to interpret the chart**  
                    
            - High 1st Win Rate (%)  
                Vendor is highly competitive and often offers the best commercial terms.  
            - High 2nd Win Rate (%)  
                Vendor consistently performs well, often just slightly less competitive than the winner.  
            - Large Gap Between 1st & 2nd Win Rate  
                Shows clear market dominance by certain vendors.
        ''')

with tab2:
    st.image("assets/2.png")
    with st.expander("See explanation"):
        st.caption('''
            The chart above shows the average price difference between 
            the lowest and second-lowest bids for each vendor when they 
            rank 1st, indicating their pricing dominance or competitiveness.
                    
            **💡 How to interpret the chart**  
                    
            - High Gap  
                High gap indicates strong vendor dominance (much lower prices).  
            - Low Gap  
                Low gap indicates intense competition with similar pricing among vendors.  
            
            The dashed line represents the average gap across all vendors, serving as a benchmark (16.0%).
        ''')
    
st.write("")
st.markdown("**:blue-badge[5. SUPER BUTTON]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Lastly, there is a 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Super Button</span> 
            feature where all dataframes generated by the system can be downloaded as a single file with multiple sheets. 
            You can also customize the order of the sheets. The interface looks more or less like this.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["Original Data", "Transpose Data"])

with tab1:
    dataframes = {
        "Merge Data": df_merge,
        "TCO Summary": df_tco_ori,
        "Bid & Price Analysis": df_analysis,
    }

    # Tampilkan multiselect
    selected_sheets = tab1.multiselect(
        "Select sheets to download in a single Excel file:",
        options=list(dataframes.keys()),
        default=list(dataframes.keys())  # default semua dipilih
    )

    # Fungsi "Super Button" & Formatting
    def generate_multi_sheet_excel(selected_sheets, df_dict):
        """
        Buat Excel multi-sheet dengan highlight:
        - Sheet 'Bid & Price Analysis' -> highlight 1st & 2nd vendor
        - Sheet lainnya -> highlight row TOTAL
        """
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for sheet in selected_sheets:
                df = df_dict[sheet].copy()
                df.to_excel(writer, index=False, sheet_name=sheet)
                workbook  = writer.book
                worksheet = writer.sheets[sheet]

                # --- Format umum ---
                fmt_rupiah = workbook.add_format({'num_format': '#,##0'})
                fmt_pct    = workbook.add_format({'num_format': '#,##0.0"%"'})
                fmt_total  = workbook.add_format({
                    "bold": True, "bg_color": "#D9EAD3", "font_color": "#1A5E20", "num_format": "#,##0"
                })
                fmt_first  = workbook.add_format({'bg_color': '#C6EFCE', "num_format": "#,##0"})
                fmt_second = workbook.add_format({'bg_color': '#FFEB9C', "num_format": "#,##0"})

                # Identifikasi numeric columns
                numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
                vendor_cols = [c for c in numeric_cols] if sheet == "Bid & Price Analysis" else []

                # Apply format kolom numeric / persen
                for col_idx, col_name in enumerate(df.columns):
                    if col_name in numeric_cols:
                        worksheet.set_column(col_idx, col_idx, 15, fmt_rupiah)
                    if "%" in col_name:
                        worksheet.set_column(col_idx, col_idx, 15, fmt_pct)

                # --- Highlight baris ---
                for row_idx, row in enumerate(df.itertuples(index=False), start=1):
                    # Cek apakah TOTAL
                    is_total_row = any(str(x).strip().upper() == "TOTAL" for x in row if pd.notna(x))

                    # Ambil nama 1st & 2nd vendor untuk sheet Bid & Price Analysis
                    if sheet == "Bid & Price Analysis":
                        first_vendor_name = row[df.columns.get_loc("1st Vendor")]
                        second_vendor_name = row[df.columns.get_loc("2nd Vendor")]

                        # Cari index kolom vendor di vendor_cols
                        first_idx = df.columns.get_loc(first_vendor_name) if first_vendor_name in vendor_cols else None
                        second_idx = df.columns.get_loc(second_vendor_name) if second_vendor_name in vendor_cols else None

                    # Loop tiap kolom
                    for col_idx, col_name in enumerate(df.columns):
                        value = row[col_idx]
                        fmt = None

                        # Highlight TOTAL untuk sheet selain Bid & Price Analysis
                        if is_total_row and sheet in ["Merge Data", "TCO Summary"]:
                            fmt = fmt_total

                        # Highlight 1st/2nd vendor
                        elif sheet == "Bid & Price Analysis":
                            if first_idx is not None and col_idx == first_idx:
                                fmt = fmt_first
                            elif second_idx is not None and col_idx == second_idx:
                                fmt = fmt_second

                        # Tangani NaN / None / inf
                        if pd.isna(value) or (isinstance(value, (int, float)) and np.isinf(value)):
                            value = ""

                        worksheet.write(row_idx, col_idx, value, fmt)

        output.seek(0)
        return output

    # ---- DOWNLOAD BUTTON ----
    if selected_sheets:
        excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

        st.download_button(
            label="Download",
            data=excel_bytes,
            file_name="super botton - original.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )


with tab2:
    dataframes = {
        "Merge Transposed": df_merge_transpose,
        "TCO Summary Transposed": df_tco_transpose,
        "Bid & Price Analysis Transposed": df_analysis_transpose,
    }

    # Tampilkan multiselect
    selected_sheets = tab2.multiselect(
        "Select sheets to download in a single Excel file:",
        options=list(dataframes.keys()),
        default=list(dataframes.keys())  # default semua dipilih
    )

    # Fungsi "Super Button" & Formatting
    def generate_multi_sheet_excel_transposed(selected_sheets, df_dict):
        """
        Buat Excel multi-sheet dengan highlight:
        - Sheet 'Bid & Price Analysis' -> highlight 1st & 2nd vendor
        - Sheet lainnya -> highlight row TOTAL
        """
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for sheet in selected_sheets:
                df = df_dict[sheet].copy()
                df.to_excel(writer, index=False, sheet_name=sheet)
                workbook  = writer.book
                worksheet = writer.sheets[sheet]

                # --- Format umum ---
                fmt_rupiah = workbook.add_format({'num_format': '#,##0'})
                fmt_pct    = workbook.add_format({'num_format': '#,##0.0"%"'})
                fmt_total  = workbook.add_format({
                    "bold": True, "bg_color": "#D9EAD3", "font_color": "#1A5E20", "num_format": "#,##0"
                })
                fmt_first  = workbook.add_format({'bg_color': '#C6EFCE', "num_format": "#,##0"})
                fmt_second = workbook.add_format({'bg_color': '#FFEB9C', "num_format": "#,##0"})

                # Identifikasi numeric columns
                numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
                vendor_cols = [c for c in numeric_cols] if sheet == "Bid & Price Analysis Transposed" else []

                # Apply format kolom numeric / persen
                for col_idx, col_name in enumerate(df.columns):
                    if col_name in numeric_cols:
                        worksheet.set_column(col_idx, col_idx, 15, fmt_rupiah)
                    if "%" in col_name:
                        worksheet.set_column(col_idx, col_idx, 15, fmt_pct)

                # --- Highlight baris ---
                for row_idx, row in enumerate(df.itertuples(index=False), start=1):
                    # Cek apakah TOTAL
                    is_total_row = any(str(x).strip().upper() == "TOTAL" for x in row if pd.notna(x))

                    # Ambil nama 1st & 2nd vendor untuk sheet Bid & Price Analysis Transposed
                    if sheet == "Bid & Price Analysis Transposed":
                        first_vendor_name = row[df.columns.get_loc("1st Vendor")]
                        second_vendor_name = row[df.columns.get_loc("2nd Vendor")]

                        # Cari index kolom vendor di vendor_cols
                        first_idx = df.columns.get_loc(first_vendor_name) if first_vendor_name in vendor_cols else None
                        second_idx = df.columns.get_loc(second_vendor_name) if second_vendor_name in vendor_cols else None

                    # Loop tiap kolom
                    for col_idx, col_name in enumerate(df.columns):
                        value = row[col_idx]
                        fmt = None

                        # Highlight TOTAL untuk sheet selain Bid & Price Analysis Transposed
                        if is_total_row and sheet in ["Merge Transposed", "TCO Summary Transposed"]:
                            fmt = fmt_total

                        # Highlight 1st/2nd vendor
                        elif sheet == "Bid & Price Analysis Transposed":
                            if first_idx is not None and col_idx == first_idx:
                                fmt = fmt_first
                            elif second_idx is not None and col_idx == second_idx:
                                fmt = fmt_second

                        # Tangani NaN / None / inf
                        if pd.isna(value) or (isinstance(value, (int, float)) and np.isinf(value)):
                            value = ""

                        worksheet.write(row_idx, col_idx, value, fmt)

        output.seek(0)
        return output

    # ---- DOWNLOAD BUTTON ----
    if selected_sheets:
        excel_bytes = generate_multi_sheet_excel_transposed(selected_sheets, dataframes)

        tab2.download_button(
            label="Download",
            data=excel_bytes,
            file_name="super botton - transpose.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

st.write("")
st.divider()

st.markdown("#### Video Tutorial")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            I have also included a video tutorial, which you can access through the 
            <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">YouTube</span> 
            link below.
        </div>
    """,
    unsafe_allow_html=True
)

st.video("https://youtu.be/2oo8SNo39A8?si=ioHAlGCnwq13Dj4V")