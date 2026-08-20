import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime

# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Task Progress Dashboard",
    page_icon="📊",
    layout="wide",
)

EXCEL_FILE = Path("tasks.xlsx")
SHEET_NAME = "Tasks"

REQUIRED_COLUMNS = [
    "Date",
    "Task/ActionItem",
    "Responsible",
    "Area/Locations",
    "Department",
    "Priority",
    "Due Date",
    "Status",
]


# ============================================================
# STYLING
# ============================================================

st.markdown(
    """
    <style>

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }

    .header {
        background: linear-gradient(
            135deg,
            #0f172a,
            #155e75
        );
        padding: 25px 30px;
        border-radius: 18px;
        color: white;
        margin-bottom: 25px;
    }

    .header h1 {
        margin: 0;
        font-size: 32px;
    }

    .header p {
        margin-top: 8px;
        margin-bottom: 0;
        opacity: 0.85;
    }

    .progress-container {
        background: #e5e7eb;
        border-radius: 20px;
        height: 30px;
        width: 100%;
        overflow: hidden;
        margin-top: 20px;
    }

    .progress-bar {
        height: 100%;
        border-radius: 20px;
        background: linear-gradient(
            90deg,
            #22c55e,
            #06b6d4
        );
        transition: width 0.5s ease;
    }

    .progress-text {
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        margin-top: 10px;
    }

    .section-title {
        font-size: 24px;
        font-weight: 700;
        margin-top: 20px;
        margin-bottom: 15px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# LOAD EXCEL
# ============================================================

@st.cache_data
def load_excel(file_path, modified_time):

    df = pd.read_excel(
        file_path,
        sheet_name=SHEET_NAME if SHEET_NAME else 0
    )

    # Clean column names
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    missing_columns = [
        column
        for column in REQUIRED_COLUMNS
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing columns in Excel: "
            + ", ".join(missing_columns)
        )

    df = df[REQUIRED_COLUMNS].copy()

    # Dates
    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    df["Due Date"] = pd.to_datetime(
        df["Due Date"],
        errors="coerce"
    )

    # Text fields
    text_columns = [
        "Task/ActionItem",
        "Responsible",
        "Area/Locations",
        "Department",
        "Priority",
        "Status",
    ]

    for column in text_columns:
        df[column] = (
            df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # Stable row identifier
    df["_row_id"] = df.index

    return df


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(df):

    save_df = df.copy()

    if "_row_id" in save_df.columns:
        save_df = save_df.drop(columns=["_row_id"])

    save_df.to_excel(
        EXCEL_FILE,
        sheet_name=SHEET_NAME,
        index=False
    )

    # Clear cached Excel data
    load_excel.clear()


# ============================================================
# DATE FORMAT
# ============================================================

def format_date(value):

    if pd.isna(value):
        return ""

    return pd.Timestamp(value).strftime("%d-%b-%Y")


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="header">
        <h1>📊 Task Progress Dashboard</h1>
        <p>
            Excel-backed task management and progress monitoring
        </p>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# CHECK EXCEL
# ============================================================

if not EXCEL_FILE.exists():

    st.error(
        f"Excel file '{EXCEL_FILE}' was not found."
    )

    st.info(
        "Place your Excel file in the same folder as "
        "app.py and name it 'tasks.xlsx'."
    )

    st.stop()


# ============================================================
# READ DATA
# ============================================================

try:

    df = load_excel(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime
    )

except Exception as error:

    st.error(
        f"Unable to read Excel file: {error}"
    )

    st.stop()


# ============================================================
# CALCULATE PROGRESS
# ============================================================

total_tasks = len(df)

completed_tasks = (
    df["Status"]
    .str.casefold()
    .eq("close")
    .sum()
)

open_tasks = total_tasks - completed_tasks

if total_tasks > 0:
    completion_percentage = (
        completed_tasks / total_tasks
    ) * 100
else:
    completion_percentage = 0


# ============================================================
# TOP SECTION
# ============================================================

left_column, right_column = st.columns(
    [1.25, 1],
    gap="large"
)


# ============================================================
# LEFT - PROGRESS
# ============================================================

with left_column:

    st.markdown(
        "### 📈 Overall Completion"
    )

    st.markdown(
        f"""
        <div class="progress-container">
            <div
                class="progress-bar"
                style="width:{completion_percentage:.1f}%">
            </div>
        </div>

        <div class="progress-text">
            {completion_percentage:.1f}% Complete
        </div>
        """,
        unsafe_allow_html=True
    )

    metric1, metric2, metric3 = st.columns(3)

    metric1.metric(
        "Total Tasks",
        total_tasks
    )

    metric2.metric(
        "Open Tasks",
        open_tasks
    )

    metric3.metric(
        "Completed",
        completed_tasks
    )


# ============================================================
# RIGHT - DEPARTMENT SUMMARY
# ============================================================

with right_column:

    st.markdown(
        "### 🏢 Department-wise Tasks"
    )

    department_data = df.copy()

    department_data["Completed"] = (
        department_data["Status"]
        .str.casefold()
        .eq("close")
    )

    department_summary = (
        department_data
        .groupby("Department", dropna=False)
        .agg(
            Tasks=("Task/ActionItem", "count"),
            Completed=("Completed", "sum")
        )
        .reset_index()
    )

    department_summary["Open"] = (
        department_summary["Tasks"]
        - department_summary["Completed"]
    )

    department_summary["Completion %"] = (
        department_summary["Completed"]
        / department_summary["Tasks"]
        * 100
    ).round(1)

    st.dataframe(
        department_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Completion %": st.column_config.ProgressColumn(
                "Completion %",
                min_value=0,
                max_value=100,
                format="%.1f%%"
            )
        }
    )


# ============================================================
# SEPARATOR
# ============================================================

st.divider()


# ============================================================
# COMPLETED TASK BUTTON
# ============================================================

if "show_completed" not in st.session_state:
    st.session_state.show_completed = False


button_text = (
    "⬆️ Hide Completed Tasks"
    if st.session_state.show_completed
    else "✅ View All Completed Tasks"
)


if st.button(button_text):

    st.session_state.show_completed = (
        not st.session_state.show_completed
    )

    st.rerun()


# ============================================================
# COMPLETED TASKS
# ============================================================

if st.session_state.show_completed:

    st.markdown(
        "### ✅ Completed Tasks"
    )

    completed_df = df[
        df["Status"]
        .str.casefold()
        .eq("close")
    ].copy()

    if completed_df.empty:

        st.info(
            "There are no completed tasks yet."
        )

    else:

        completed_df["Date"] = (
            completed_df["Date"]
            .apply(format_date)
        )

        completed_df["Due Date"] = (
            completed_df["Due Date"]
            .apply(format_date)
        )

        completed_display = completed_df[
            [
                "Date",
                "Task/ActionItem",
                "Responsible",
                "Area/Locations",
                "Department",
                "Priority",
                "Due Date",
                "Status",
            ]
        ].rename(
            columns={
                "Task/ActionItem": "Task",
                "Area/Locations": "Location",
            }
        )

        st.dataframe(
            completed_display,
            use_container_width=True,
            hide_index=True
        )

    st.divider()


# ============================================================
# OPEN TASKS
# ============================================================

st.markdown(
    "### 📋 Open Tasks"
)


open_df = df[
    ~df["Status"]
    .str.casefold()
    .eq("close")
].copy()


# ============================================================
# FILTERS
# ============================================================

with st.expander(
    "🔎 Search & Filter Open Tasks",
    expanded=True
):

    filter1, filter2, filter3, filter4 = st.columns(4)

    search_text = filter1.text_input(
        "Search",
        placeholder=(
            "Task, responsible, location..."
        )
    )

    department_options = sorted(
        [
            value
            for value in open_df["Department"].unique()
            if value
        ]
    )

    priority_options = sorted(
        [
            value
            for value in open_df["Priority"].unique()
            if value
        ]
    )

    status_options = sorted(
        [
            value
            for value in open_df["Status"].unique()
            if value
        ]
    )

    selected_departments = (
        filter2.multiselect(
            "Department",
            department_options
        )
    )

    selected_priorities = (
        filter3.multiselect(
            "Priority",
            priority_options
        )
    )

    selected_statuses = (
        filter4.multiselect(
            "Status",
            status_options
        )
    )


# ============================================================
# APPLY SEARCH
# ============================================================

if search_text:

    search_value = search_text.casefold()

    search_mask = open_df.apply(
        lambda row:
        search_value in
        " ".join(
            row.astype(str)
        ).casefold(),
        axis=1
    )

    open_df = open_df[
        search_mask
    ]


# ============================================================
# APPLY DEPARTMENT FILTER
# ============================================================

if selected_departments:

    open_df = open_df[
        open_df["Department"]
        .isin(selected_departments)
    ]


# ============================================================
# APPLY PRIORITY FILTER
# ============================================================

if selected_priorities:

    open_df = open_df[
        open_df["Priority"]
        .isin(selected_priorities)
    ]


# ============================================================
# APPLY STATUS FILTER
# ============================================================

if selected_statuses:

    open_df = open_df[
        open_df["Status"]
        .isin(selected_statuses)
    ]


st.caption(
    f"Showing {len(open_df)} open task(s)."
)


# ============================================================
# EDITABLE TASK TABLE
# ============================================================

if not open_df.empty:

    table_data = open_df[
        [
            "_row_id",
            "Task/ActionItem",
            "Responsible",
            "Area/Locations",
            "Department",
            "Priority",
            "Due Date",
            "Status",
        ]
    ].copy()

    table_data["Due Date"] = pd.to_datetime(
        table_data["Due Date"],
        errors="coerce"
    ).dt.date

    editable_data = table_data.drop(
        columns=["_row_id"]
    )

    edited_data = st.data_editor(
        editable_data,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="task_editor",

        column_config={

            "Task/ActionItem":
                st.column_config.TextColumn(
                    "Task",
                    required=True
                ),

            "Responsible":
                st.column_config.TextColumn(
                    "Responsible"
                ),

            "Area/Locations":
                st.column_config.TextColumn(
                    "Location"
                ),

            "Department":
                st.column_config.TextColumn(
                    "Department"
                ),

            "Priority":
                st.column_config.SelectboxColumn(
                    "Priority",
                    options=sorted(
                        set(
                            priority_options
                            + [
                                "Low",
                                "Medium",
                                "High",
                                "Critical"
                            ]
                        )
                    )
                ),

            "Due Date":
                st.column_config.DateColumn(
                    "Due Date"
                ),

            "Status":
                st.column_config.SelectboxColumn(
                    "Status",
                    options=sorted(
                        set(
                            status_options
                            + [
                                "Open",
                                "In Progress",
                                "Close"
                            ]
                        )
                    )
                ),
        }
    )


    # ========================================================
    # UPDATE BUTTON
    # ========================================================

    if st.button(
        "💾 Update Excel",
        type="primary"
    ):

        updated_df = df.copy()

        for position in range(
            len(edited_data)
        ):

            original_row_id = (
                open_df.iloc[position]["_row_id"]
            )

            row = edited_data.iloc[position]

            updated_df.loc[
                original_row_id,
                "Task/ActionItem"
            ] = row["Task/ActionItem"]

            updated_df.loc[
                original_row_id,
                "Responsible"
            ] = row["Responsible"]

            updated_df.loc[
                original_row_id,
                "Area/Locations"
            ] = row["Area/Locations"]

            updated_df.loc[
                original_row_id,
                "Department"
            ] = row["Department"]

            updated_df.loc[
                original_row_id,
                "Priority"
            ] = row["Priority"]

            updated_df.loc[
                original_row_id,
                "Due Date"
            ] = pd.to_datetime(
                row["Due Date"],
                errors="coerce"
            )

            updated_df.loc[
                original_row_id,
                "Status"
            ] = row["Status"]


        # Save to Excel
        save_excel(updated_df)

        st.success(
            "✅ Excel updated successfully."
        )

        # Refresh all sections
        st.rerun()


else:

    st.success(
        "🎉 No open tasks match the current filters."
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "A task is considered completed only when "
    "Status = Close."
)

st.caption(
    "Last dashboard refresh: "
    + datetime.now().strftime(
        "%d-%b-%Y %H:%M:%S"
    )
)
