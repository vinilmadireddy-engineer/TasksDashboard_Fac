import time
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="GR/FCM-Ban OPL Points",
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
    "Comments",
    "Completion Status"
]

PRIORITY_VALUES = ["Low", "Medium", "High", "Critical"]
STATUS_VALUES = ["Open", "In Progress", "Close"]
COMPLETION_STATUS_VALUES = [0, 10, 20, 30, 40, 50, 70, 80, 90, 100]

# ============================================================
# STYLING WITH SINGLE COLOR BACKGROUND
# ============================================================

st.markdown(
    """
    <style>
    html, body {
        height: 100%;
        margin: 0 !important;
        padding: 0 !important;
    }

    /* Solid background color */
    .stApp {
        background: #007bc0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    .table {
        font-size: 18px !important;
        font-weight: bold !important;
        background-color: rgb(14, 17, 23);
    }

    .table th {
        font-size: 20px !important;
        font-weight: bold !important;
        background-color: rgb(14, 17, 23)
    }

    /* Keep containers transparent so background shows */
    .block-container {
        background: transparent !important;
        position: relative;
        z-index: 1;
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }

    /* Preserve your original styles */
    .header {
        background: linear-gradient(135deg, #0f172a, #155e75);
        padding: 25px 30px;
        border-radius: 18px;
        color: white;
        margin-bottom: 20px;
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
        background: linear-gradient(90deg, #22c55e, #06b6d4);
        transition: width 0.5s ease;
    }

    .progress-text {
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        margin-top: 10px;
        color: #111827;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# SESSION STATE
# ============================================================

if "last_save" not in st.session_state:
    st.session_state.last_save = None

if "save_message" not in st.session_state:
    st.session_state.save_message = None

if "show_completed" not in st.session_state:
    st.session_state.show_completed = False


# ============================================================
# LOAD AND SAVE EXCEL
# ============================================================

@st.cache_data
def load_excel(file_path, modified_time):
    # modified_time is included to invalidate the cache when Excel changes.
    del modified_time

    loaded_df = pd.read_excel(
        file_path,
        sheet_name=SHEET_NAME,
        engine="openpyxl",
    )

    loaded_df.columns = [str(column).strip() for column in loaded_df.columns]

    missing_columns = [
        column for column in REQUIRED_COLUMNS
        if column not in loaded_df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing columns in Excel: " + ", ".join(missing_columns)
        )

    loaded_df = loaded_df[REQUIRED_COLUMNS].copy()
    loaded_df["Date"] = pd.to_datetime(loaded_df["Date"], errors="coerce")
    loaded_df["Due Date"] = pd.to_datetime(
        loaded_df["Due Date"], errors="coerce"
    )

    text_columns = [
        "Task/ActionItem",
        "Responsible",
        "Area/Locations",
        "Department",
        "Priority",
        "Status",
    ]
    loaded_df["Completion Status"] = (
        pd.to_numeric(loaded_df["Completion Status"], errors="coerce")
        .fillna(COMPLETION_STATUS_VALUES[0])
        .round()
        .astype(int)
    )
    loaded_df["Completion Status"] = loaded_df["Completion Status"].where(
        loaded_df["Completion Status"].isin(COMPLETION_STATUS_VALUES),
        COMPLETION_STATUS_VALUES[0],
    )

    for column in text_columns:
        loaded_df[column] = (
            loaded_df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    loaded_df["_row_id"] = loaded_df.index
    return loaded_df

# ============================================================
# IST TIME
# ============================================================

def get_ist_time():

    return datetime.now(
        ZoneInfo("Asia/Kolkata")
    )

def save_excel(dataframe):
    """Safely save tasks and verify that the written workbook is readable."""
    save_df = dataframe.copy()
    save_df = save_df.drop(columns=["_row_id"], errors="ignore")
    save_df = save_df[REQUIRED_COLUMNS]

    temporary_file = EXCEL_FILE.with_name(
        f"{EXCEL_FILE.stem}_temp.xlsx"
    )

    try:
        save_df.to_excel(
            temporary_file,
            sheet_name=SHEET_NAME,
            index=False,
            engine="openpyxl",
        )

        # Verify the temporary workbook before replacing the active file.
        verification_df = pd.read_excel(
            temporary_file,
            sheet_name=SHEET_NAME,
            engine="openpyxl",
        )

        if len(verification_df) != len(save_df):
            raise IOError(
                "Excel verification failed because the saved row count "
                "does not match the expected row count."
            )

        temporary_file.replace(EXCEL_FILE)
        load_excel.clear()

    except Exception:
        if temporary_file.exists():
            temporary_file.unlink()
        raise


def unique_non_empty_values(series):
    return sorted(
        {
            str(value).strip()
            for value in series.dropna()
            if str(value).strip()
        }
    )


def synchronize_status_and_completion(status, completion):
    clean_status = str(status).strip()
    if pd.isna(completion) or str(completion).strip() == "":
        completion_value = 0
    else:
        completion_value = int(float(completion))

    if clean_status.casefold() == "close" or completion_value == 100:
        return "Close", 100

    return clean_status, completion_value


def format_date(value):
    if pd.isna(value):
        return ""
    return pd.Timestamp(value).strftime("%d-%b-%Y")


def create_excel_download(dataframe):
    export_df = dataframe.copy()
    export_df = export_df.drop(columns=["_row_id"], errors="ignore")
    export_df = export_df[REQUIRED_COLUMNS]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(
            writer,
            sheet_name=SHEET_NAME,
            index=False,
        )
    return output.getvalue()


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="header">
        <h1>📊 GR/FCM-Ban OPL Points</h1>
        <p>Task management and progress monitoring</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# This message survives st.rerun(), so users can clearly see save success.
if st.session_state.save_message:
    st.success(st.session_state.save_message, icon="✅")

if st.session_state.last_save:
    st.caption(f"💾 Last successful Excel update: {st.session_state.last_save}")


# ============================================================
# CHECK AND READ EXCEL
# ============================================================

if not EXCEL_FILE.exists():
    st.error(f"Excel file '{EXCEL_FILE}' was not found.")
    st.info(
        "Place tasks.xlsx in the same folder as app.py. "
        "The workbook must contain a sheet named 'Tasks'."
    )
    st.stop()

try:
    df = load_excel(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime_ns,
    )
except Exception as error:
    st.error(f"Unable to read Excel file: {error}")
    st.stop()


# ============================================================
# ADD NEW TASK
# ============================================================

with st.expander("➕ Add New Task", expanded=False):
    existing_departments = unique_non_empty_values(df["Department"])

    with st.form("add_task_form", clear_on_submit=True):
        add_col1, add_col2 = st.columns(2)

        task_name = add_col1.text_input(
            "Task / Action Item *",
            placeholder="Enter the task description",
        )
        responsible = add_col2.text_input(
            "Responsible",
            placeholder="Enter responsible person's name",
        )

        add_col3, add_col4 = st.columns(2)
        location = add_col3.text_input(
            "Area / Location",
            placeholder="Enter area or location",
        )
        department_selection = add_col4.selectbox(
            "Department",
            options=["Enter new department"] + existing_departments,
        )

        if department_selection == "Enter new department":
            department = st.text_input(
                "New Department *",
                placeholder="Enter department name",
            )
        else:
            department = department_selection

        add_col5, add_col6 = st.columns(2)
        priority = add_col5.selectbox(
            "Priority",
            options=PRIORITY_VALUES,
            index=1,
        )
        due_date = add_col6.date_input(
            "Due Date",
            value=None,
            format="DD-MM-YYYY",
        )

        add_col7, add_col8 = st.columns(2)
        task_date = add_col7.date_input(
            "Task Created Date",
            value=datetime.now().date(),
            format="DD-MM-YYYY",
        )
        status = add_col8.selectbox(
            "Status",
            options=STATUS_VALUES,
            index=0,
        )
        comments = st.text_area(
            "Comments",
             placeholder="Add any notes or comments about this task",
        )


        add_task_submitted = st.form_submit_button(
            "➕ Add Task",
            type="primary",
            use_container_width=True,
        )

    if add_task_submitted:
        clean_task_name = task_name.strip()
        clean_department = department.strip()

        if not clean_task_name:
            st.error("Task / Action Item is required.")
        elif not clean_department:
            st.error("Department is required.")
        else:
            progress = st.progress(0, text="Preparing the new task...")

            try:
                progress.progress(25, text="Validating task details...")

                new_task = pd.DataFrame(
                    [{
                        "Date": pd.to_datetime(task_date),
                        "Task/ActionItem": clean_task_name,
                        "Responsible": responsible.strip(),
                        "Area/Locations": location.strip(),
                        "Department": clean_department,
                        "Priority": priority,
                        "Due Date": (
                            pd.to_datetime(due_date)
                            if due_date is not None
                            else pd.NaT
                        ),
                        "Status": status,
                        "Comments": comments.strip(),
                        "Completion Status": COMPLETION_STATUS_VALUES[0],
                    }]
                )

                updated_df = pd.concat(
                    [
                        df.drop(columns=["_row_id"], errors="ignore"),
                        new_task,
                    ],
                    ignore_index=True,
                )

                progress.progress(65, text="Writing the new task to Excel...")
                save_excel(updated_df)

                saved_at = get_ist_time().now().strftime("%d-%b-%Y %H:%M:%S")
                st.session_state.last_save = saved_at
                st.session_state.save_message = (
                    f"Task '{clean_task_name}' was added and Excel was "
                    f"updated successfully at {saved_at}."
                )

                progress.progress(100, text="Excel updated successfully.")
                time.sleep(0.6)
                st.rerun()
                

            except PermissionError:
                progress.empty()
                st.error(
                    "Unable to update tasks.xlsx. Close the workbook in "
                    "Microsoft Excel and try again."
                )
            except Exception as error:
                progress.empty()
                st.error(f"Unable to add the task: {error}")

# ============================================================
# PROGRESS SUMMARY
# ============================================================

total_tasks = len(df)
completed_tasks = int(df["Status"].str.casefold().eq("close").sum())
open_tasks = total_tasks - completed_tasks
completion_percentage = (completed_tasks / total_tasks * 100 if total_tasks else 0)

left_column, right_column = st.columns([1.25, 1], gap="large")

with left_column:
    st.markdown("### 📈 Overall Completion")
    st.markdown(
        f"""
        <div class="progress-container">
            <div class="progress-bar"
                 style="width:{completion_percentage:.1f}%"></div>
        </div>
        <div class="progress-text">
            {completion_percentage:.1f}% Complete
        </div>
        """,
        unsafe_allow_html=True,
    )

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Total Tasks", total_tasks)
    metric2.metric("Open Tasks", open_tasks)
    metric3.metric("Completed", completed_tasks)


with right_column:
    # ============================================================
    # PRIORITY-WISE TASKS WITH DARK BACKGROUND + BIGGER FONT
    # ============================================================

    st.markdown("### 🎯 Priority-wise Task Summary")

    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

    if df.empty:
        st.info("No tasks available yet.")
    else:
        priority_data = df.copy()
        priority_data["Completed"] = priority_data["Status"].str.casefold().eq("close")

        priority_summary = (
            priority_data
            .groupby("Priority", dropna=False)
            .agg(
                Tasks=("Task/ActionItem", "count"),
                Completed=("Completed", "sum"),
            )
            .reset_index()
        )
        priority_summary["Open"] = priority_summary["Tasks"] - priority_summary["Completed"]

        # ✅ Build custom HTML table with dark background + larger font
        table_html = """
        <style>
        table.priority-table {
            width: 100%;
            border-collapse: collapse;
            background-color: rgb(14, 17, 23);
            color: white;
            font-size: 18px;   /* Bigger font */
        }
        table.priority-table th, table.priority-table td {
            padding: 10px;
            text-align: center;
            border: 1px solid #444;
        }
        table.priority-table th {
            font-size: 20px;
            font-weight: bold;
            background-color: rgb(14, 17, 23);
        }
        </style>
        <table class="priority-table">
            <tr>
                <th>Priority</th>
                <th>Total Tasks</th>
                <th>Completed</th>
                <th>Open</th>
            </tr>
        """

        for _, row in priority_summary.iterrows():
            table_html += f"""
            <tr>
                <td>{row['Priority']}</td>
                <td>{row['Tasks']}</td>
                <td>{row['Completed']}</td>
                <td>{row['Open']}</td>
            </tr>
            """

        table_html += "</table>"

        # ✅ Render HTML properly
        st.components.v1.html(table_html, height=200, scrolling=True)


st.divider()

# ============================================================
# DEPARTMENT-WISE TASKS WITH DARK BACKGROUND + FONT + PROGRESS BAR
# ============================================================

st.markdown("### 🏢 Department-wise Tasks")

if df.empty:
    st.info("No tasks are available yet.")
else:
    department_data = df.copy()
    department_data["Completed"] = department_data["Status"].str.casefold().eq("close")

    department_summary = (
        department_data
        .groupby("Department", dropna=False)
        .agg(
            Tasks=("Task/ActionItem", "count"),
            Completed=("Completed", "sum"),
        )
        .reset_index()
    )

    department_summary["Open"] = department_summary["Tasks"] - department_summary["Completed"]
    department_summary["Completion %"] = (
        department_summary["Completed"] / department_summary["Tasks"] * 100
    ).round(1)

    # ✅ Build full HTML table string
    table_html = """
    <style>
    table.custom-table {
        width: 100%;
        border-collapse: collapse;
        background-color: rgb(14, 17, 23);
        color: white;
        font-size: 18px;   /* Bigger font */
    }
    table.custom-table th, table.custom-table td {
        padding: 10px;
        text-align: center;
        border: 1px solid #444;
    }
    table.custom-table th {
        font-size: 20px;
        font-weight: bold;
        background-color: rgb(14, 17, 23);
    }
    .progress-container {
        background: #333;
        border-radius: 10px;
        height: 20px;
        width: 100%;
        overflow: hidden;
    }
    .progress-bar {
        height: 100%;
        border-radius: 10px;
        background: linear-gradient(90deg, #22c55e, #06b6d4);
    }
    </style>
    <table class="custom-table">
        <tr>
            <th>Department</th>
            <th>Total Tasks</th>
            <th>Completed</th>
            <th>Open</th>
            <th>Completion %</th>
        </tr>
    """

    for _, row in department_summary.iterrows():
        table_html += f"""
        <tr>
            <td>{row['Department']}</td>
            <td>{row['Tasks']}</td>
            <td>{row['Completed']}</td>
            <td>{row['Open']}</td>
            <td>
                <div class="progress-container">
                    <div class="progress-bar" style="width:{row['Completion %']}%"></div>
                </div>
                {row['Completion %']}%
            </td>
        </tr>
        """

    table_html += "</table>"

    # ✅ Render HTML properly with components.html
    st.components.v1.html(table_html, height=600, scrolling=True)



# ============================================================
# COMPLETED TASKS
# ============================================================

st.divider()

button_text = (
    "⬆️ Hide Completed Tasks"
    if st.session_state.show_completed
    else "✅ View All Completed Tasks"
)

if st.button(button_text):
    st.session_state.show_completed = not st.session_state.show_completed
    st.rerun()

if st.session_state.show_completed:
    st.markdown("### ✅ Completed Tasks")

    completed_df = df[
        df["Status"].str.casefold().eq("close")
    ].copy()

    if completed_df.empty:
        st.info("There are no completed tasks yet.")
    else:
        completed_df["Date"] = completed_df["Date"].apply(format_date)
        completed_df["Due Date"] = completed_df["Due Date"].apply(
            format_date
        )

        completed_display = completed_df[
            REQUIRED_COLUMNS
        ].rename(
            columns={
                "Task/ActionItem": "Task",
                "Area/Locations": "Location",
            }
        )

        st.dataframe(
            completed_display,
            use_container_width=True,
            hide_index=True,
        )

    st.divider()


# ============================================================
# OPEN TASKS AND FILTERS
# ============================================================

st.markdown("### 📋 Open Tasks")

open_df = df[
    ~df["Status"].str.casefold().eq("close")
].copy()

with st.expander("🔎 Search & Filter Open Tasks", expanded=True):
    filter1, filter2, filter3, filter4 = st.columns(4)

    search_text = filter1.text_input(
        "Search",
        placeholder="Task, responsible, location...",
    )

    department_options = unique_non_empty_values(open_df["Department"])
    priority_options = unique_non_empty_values(open_df["Priority"])
    status_options = unique_non_empty_values(open_df["Status"])

    selected_departments = filter2.multiselect(
        "Department", department_options
    )
    selected_priorities = filter3.multiselect(
        "Priority", priority_options
    )
    selected_statuses = filter4.multiselect(
        "Status", status_options
    )

if search_text:
    search_columns = [
        "Task/ActionItem",
        "Responsible",
        "Area/Locations",
        "Department",
        "Priority",
        "Status",
    ]

    search_mask = (
        open_df[search_columns]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .str.casefold()
        .str.contains(search_text.casefold(), regex=False)
    )
    open_df = open_df[search_mask]

if selected_departments:
    open_df = open_df[open_df["Department"].isin(selected_departments)]
if selected_priorities:
    open_df = open_df[open_df["Priority"].isin(selected_priorities)]
if selected_statuses:
    open_df = open_df[open_df["Status"].isin(selected_statuses)]

st.caption(f"Showing {len(open_df)} open task(s).")


# ============================================================
# EDITABLE TASK TABLE AND UPDATE EXCEL
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
            "Comments",
            "Completion Status",
        ]
    ].copy()

    table_data["Due Date"] = pd.to_datetime(
        table_data["Due Date"], errors="coerce"
    ).dt.date

    # ✅ Force Comments to string to avoid type mismatch
    table_data["Comments"] = table_data["Comments"].astype(str)
    table_data["Completion Status"] = pd.to_numeric(
        table_data["Completion Status"], errors="coerce"
    ).fillna(COMPLETION_STATUS_VALUES[0]).astype(int)

    editable_data = table_data.drop(columns=["_row_id"])

    edited_data = st.data_editor(
        editable_data,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="task_editor",
        column_config={
            "Task/ActionItem": st.column_config.TextColumn("Task", required=True),
            "Responsible": st.column_config.TextColumn("Responsible"),
            "Area/Locations": st.column_config.TextColumn("Location"),
            "Department": st.column_config.TextColumn("Department"),
            "Priority": st.column_config.SelectboxColumn(
                "Priority",
                options=sorted(set(priority_options + PRIORITY_VALUES)),
            ),
            "Due Date": st.column_config.DateColumn("Due Date", format="DD-MM-YYYY"),
            "Status": st.column_config.SelectboxColumn(
                "Status",
                options=sorted(set(status_options + STATUS_VALUES)),
            ),
            "Comments": st.column_config.TextColumn("Comments"),  # ✅ safe now
            "Completion Status": st.column_config.SelectboxColumn(
                "Completion Status",
                options=COMPLETION_STATUS_VALUES,
            ),
        },
    )

    for position in range(len(edited_data)):
        synchronized_status, synchronized_completion = (
            synchronize_status_and_completion(
                edited_data.iloc[position]["Status"],
                edited_data.iloc[position]["Completion Status"],
            )
        )
        edited_data.iloc[position, edited_data.columns.get_loc("Status")] = (
            synchronized_status
        )
        edited_data.iloc[
            position,
            edited_data.columns.get_loc("Completion Status"),
        ] = synchronized_completion

    st.markdown(
    """
    <style>
    /* Target the Completion Status column cells */
    div[data-testid="stDataFrame"] td:nth-child(9) {
        position: relative;
    }
    div[data-testid="stDataFrame"] td:nth-child(9)::before {
        content: "";
        position: absolute;
        top: 2px;
        left: 2px;
        height: 80%;
        background: linear-gradient(90deg, #22c55e, #06b6d4);
        border-radius: 6px;
        z-index: 0;
        width: calc(var(--cellValue) * 1%);
    }
    div[data-testid="stDataFrame"] td:nth-child(9) span {
        position: relative;
        z-index: 1;
        color: white;
        font-weight: bold;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


    if st.button("💾 Update Excel", type="primary"):
        blank_tasks = (
            edited_data["Task/ActionItem"]
            .fillna("")
            .astype(str)
            .str.strip()
            .eq("")
        )

        if blank_tasks.any():
            st.error("Task / Action Item cannot be blank.")
        else:
            progress = st.progress(0, text="Preparing task updates...")
            status_placeholder = st.empty()

            try:
                updated_df = df.copy()
                total_rows = len(edited_data)

                status_placeholder.info("Validating edited task data...")
                progress.progress(15, text="Validating edited task data...")

                for position in range(total_rows):
                    original_row_id = int(
                        table_data.iloc[position]["_row_id"]
                    )
                    row = edited_data.iloc[position]

                    updated_df.loc[
                        original_row_id, "Task/ActionItem"
                    ] = str(row["Task/ActionItem"]).strip()
                    updated_df.loc[
                        original_row_id, "Responsible"
                    ] = str(row["Responsible"]).strip()
                    updated_df.loc[
                        original_row_id, "Area/Locations"
                    ] = str(row["Area/Locations"]).strip()
                    updated_df.loc[
                        original_row_id, "Department"
                    ] = str(row["Department"]).strip()
                    updated_df.loc[
                        original_row_id, "Priority"
                    ] = str(row["Priority"]).strip()
                    updated_df.loc[
                        original_row_id, "Due Date"
                    ] = pd.to_datetime(row["Due Date"], errors="coerce")
                    synchronized_status, synchronized_completion = (
                        synchronize_status_and_completion(
                            row["Status"], row["Completion Status"]
                        )
                    )
                    updated_df.loc[
                        original_row_id, "Status"
                    ] = synchronized_status
                    updated_df.loc[
                        original_row_id, "Comments"
                    ] = str(row["Comments"]).strip()
                    updated_df.loc[
                        original_row_id, "Completion Status"
                    ] = synchronized_completion
                    row_progress = 15 + int(
                        ((position + 1) / total_rows) * 45
                    )
                    progress.progress(
                        row_progress,
                        text=f"Preparing row {position + 1} of {total_rows}...",
                    )

                status_placeholder.info("Writing changes to Excel...")
                progress.progress(70, text="Writing changes to Excel...")

                save_excel(updated_df)

                progress.progress(90, text="Verifying the updated workbook...")

                # Read the saved workbook again to confirm it is accessible.
                verified_df = pd.read_excel(
                    EXCEL_FILE,
                    sheet_name=SHEET_NAME,
                    engine="openpyxl",
                )

                if len(verified_df) != len(updated_df):
                    raise IOError(
                        "Saved workbook verification failed."
                    )

                saved_at = get_ist_time().now().strftime(
                    "%d-%b-%Y %H:%M:%S"
                )

                st.session_state.last_save = saved_at

                progress.progress(
                    100,
                    text="Excel updated successfully."
                )

                status_placeholder.success(
                    f"✅ Excel updated successfully at {saved_at}",
                    icon="✅"
                )

                # Prepare download
                excel_bytes = create_excel_download(updated_df)

                st.download_button(
                    label="📥 Download Updated Excel",
                    data=excel_bytes,
                    file_name=(
                        "tasks_updated_"
                        + get_ist_time().now().strftime("%Y%m%d_%H%M%S")
                        + ".xlsx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )

            except PermissionError:
                progress.empty()
                status_placeholder.error(
                    "Update failed. Close tasks.xlsx in Microsoft Excel "
                    "and try again."
                )
            except Exception as error:
                progress.empty()
                status_placeholder.error(
                    f"Excel update failed: {error}"
                )
else:
    st.success("🎉 No open tasks match the current filters.")


# ============================================================
# DOWNLOAD UPDATED EXCEL
# ============================================================




# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption("A task is considered completed only when Status = Close.")
st.caption(
    "Last dashboard refresh: "
    + get_ist_time().now().strftime("%d-%b-%Y %H:%M:%S")
)
