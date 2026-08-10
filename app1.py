import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Inspecting Officer Organizer | SSC",
    layout="wide"
)

st.markdown("""
<style>
.nic-header{
    background:#7b1e1e;
    color:white;
    padding:16px;
    font-size:26px;
    font-weight:700;
}
.nic-sub{
    background:white;
    padding:10px;
    border-bottom:3px solid #7b1e1e;
}
.nic-card{
    background:white;
    padding:18px;
    border:1px solid #ccc;
    border-radius:8px;
    margin-bottom:20px;
}
.nic-title{
    font-size:18px;
    font-weight:bold;
    color:#7b1e1e;
    border-bottom:2px solid #7b1e1e;
    padding-bottom:6px;
    margin-bottom:12px;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="nic-header">Inspecting Officer Organizer</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="nic-sub">Staff Selection Commission | Examination Duty Management System</div>',
    unsafe_allow_html=True
)

EXAMS = [
    "Combined Graduate Level Examination",
    "Combined Higher Secondary Examination",
    "Junior Engineer Examination",
    "Selection Post Examination"
]

with st.sidebar:

    st.markdown("### 🛠 Control Panel")

    selected_exam = st.selectbox(
        "Examination",
        EXAMS
    )

    exam_year = st.selectbox(
        "Year",
        range(2025,2037)
    )

    custom_exam_name = st.text_input(
        "Custom Examination Name"
    )

    st.markdown("---")

    uploaded_file = st.file_uploader(
        "Upload Inspecting Officers Master Excel",
        type=["xlsx"]
    )

if uploaded_file is None:
    st.info("Please upload Excel file.")
    st.stop()

final_exam_name = (
    custom_exam_name.strip()
    if custom_exam_name
    else selected_exam
)

try:
    df = pd.read_excel(uploaded_file)

except Exception as e:
    st.error(f"Error reading Excel : {e}")
    st.stop()

df.columns = df.columns.str.strip()

NAME_COL="Name of Inspecting Officer"
GROUP_COL="GROUP 'A' & 'B'"
STATUS_COL="status"
BANK_COL="Name of Bank & Branch"
ACC_COL="Account No"
IFSC_COL="IFSC CODE"
MOBILE_COL="Mob. No."
CITY_COL="CITY"

required_cols=[
    NAME_COL,
    GROUP_COL,
    STATUS_COL,
    BANK_COL,
    ACC_COL,
    IFSC_COL,
    CITY_COL
]

missing=[c for c in required_cols if c not in df.columns]

if missing:
    st.error(f"Missing Columns : {missing}")
    st.stop()

if MOBILE_COL not in df.columns:
    df[MOBILE_COL]=""

df[STATUS_COL]=(
    df[STATUS_COL]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.lower()
)

df=df.dropna(how="all")

df=df[
    df[NAME_COL]
    .astype(str)
    .str.strip()!=""
]

df=df.reset_index(drop=True)
def remuneration(group, full_days, single_days, status):

    full_rate = 0
    single_rate = 0

    if status == "retired":

        if group == "A":
            full_rate = 2000
            single_rate = 1200

        elif group == "B":
            full_rate = 1500
            single_rate = 900

    elif status == "working":

        if group == "A":
            full_rate = 1200
            single_rate = 1000

        elif group == "B":
            full_rate = 900
            single_rate = 750

    return (full_days * full_rate) + (single_days * single_rate)


def distance_allowance(km):

    if km <= 20:
        return 300

    elif km <= 50:
        return 500

    return 750


def create_bordered_excel(df, exam_name, exam_year):

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        start_row = 3

        df.to_excel(
            writer,
            index=False,
            sheet_name="Final Summary",
            startrow=start_row
        )

        ws = writer.sheets["Final Summary"]

        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=df.shape[1]
        )

        title = ws.cell(1, 1)
        title.value = f"{exam_name} - {exam_year}"
        title.font = Font(size=14, bold=True)
        title.alignment = Alignment(horizontal="center")

        thin = Side(style="thin")

        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        for row in ws.iter_rows():

            for cell in row:

                cell.border = border
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                if cell.row == start_row + 1:
                    cell.font = Font(bold=True)

        for col in range(1, ws.max_column + 1):

            ws.column_dimensions[
                get_column_letter(col)
            ].width = 22

    output.seek(0)

    return output


if "assignments" not in st.session_state:
    st.session_state.assignments = []

if "officer_distance" not in st.session_state:
    st.session_state.officer_distance = {}
    st.markdown('<div class="nic-card">', unsafe_allow_html=True)

st.markdown(
    '<div class="nic-title">Select City & Inspecting Officers</div>',
    unsafe_allow_html=True
)

cities = sorted(df[CITY_COL].dropna().unique())

selected_city = st.selectbox(
    "City",
    cities
)

city_df = df[df[CITY_COL] == selected_city]

selected_officers = st.multiselect(
    "Inspecting Officers",
    sorted(city_df[NAME_COL].unique())
)

st.markdown("</div>", unsafe_allow_html=True)

if selected_officers:

    st.markdown('<div class="nic-card">', unsafe_allow_html=True)

    st.markdown(
        '<div class="nic-title">Officer-wise Assignment</div>',
        unsafe_allow_html=True
    )

    for officer in selected_officers:

        row = city_df[
            city_df[NAME_COL] == officer
        ].iloc[0]

        group = str(row[GROUP_COL]).strip()
        status = str(row[STATUS_COL]).strip()

        with st.expander(
            f"👤 {officer} | Group {group} | {status.capitalize()}"
        ):

            distance = st.number_input(
                "Distance from HQ (KM)",
                min_value=0,
                value=0,
                key=f"dist_{officer}"
            )

            full_days = st.number_input(
                "Full Shift Days",
                min_value=0,
                value=0,
                key=f"full_{officer}"
            )

            single_days = st.number_input(
                "Single Shift Days",
                min_value=0,
                value=0,
                key=f"single_{officer}"
            )

            if st.button(
                "Add / Update",
                key=f"add_{officer}"
            ):

                st.session_state.assignments = [
                    a for a in st.session_state.assignments
                    if a["Officer"] != officer
                ]

                da = distance_allowance(distance)

                total_days = full_days + single_days

                st.session_state.assignments.append({

                    "Officer": officer,

                    "City": row[CITY_COL],

                    "Group": group,

                    "Status": status.capitalize(),

                    "Bank Name & Branch": row[BANK_COL],

                    "Account No": row[ACC_COL],

                    "IFSC Code": row[IFSC_COL],

                    "Mobile No": row.get(MOBILE_COL, ""),

                    "Distance (KM)": distance,

                    "Full Shift Days": full_days,

                    "Single Shift Days": single_days,

                    "Total Days": total_days,

                    "Remuneration": remuneration(
                        group,
                        full_days,
                        single_days,
                        status
                    ),

                    "Distance Allowance": da * total_days

                })

                st.success(
                    f"{officer} updated successfully."
                )

    st.markdown("</div>", unsafe_allow_html=True)
if st.session_state.assignments:

    summary = pd.DataFrame(st.session_state.assignments)

    summary["Grand Total"] = (
        summary["Remuneration"] +
        summary["Distance Allowance"]
    )

    summary.insert(
        0,
        "Sl. No.",
        range(1, len(summary) + 1)
    )

    summary = summary[
        [
            "Sl. No.",
            "Officer",
            "City",
            "Group",
            "Status",
            "Bank Name & Branch",
            "Account No",
            "IFSC Code",
            "Mobile No",
            "Distance (KM)",
            "Full Shift Days",
            "Single Shift Days",
            "Total Days",
            "Remuneration",
            "Distance Allowance",
            "Grand Total"
        ]
    ]

    st.markdown("### 📊 Final Summary")

    st.dataframe(
        summary,
        use_container_width=True
    )

    excel_file = create_bordered_excel(
        summary,
        final_exam_name,
        exam_year
    )

    st.download_button(
        "⬇ Download Final Summary (Excel)",
        data=excel_file,
        file_name=f"{final_exam_name}_{exam_year}_Final_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )