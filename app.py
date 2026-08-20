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

PRIORITY_VALUES = [
    "Low",
    "Medium",
    "High",
    "Critical",
]

STATUS_VALUES = [
    "Open",
    "In Progress",
    "Close",
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
        background: linear-gradient(135deg, #0f172a, #155e75);
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
        background: linear-gradient(90deg, #22c55e, #06b6d4);
        transition: width 0.5s ease;
    }

    .progress-text {
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        margin-top: 10px;
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
    # modified_time is intentionally passed so the cache is invalidated
    # whenever the workbook changes.
    del modified_time

    loaded_df = pd.read_excel(
        file_path,
        sheet_name=SHEET_NAME if SHEET_NAME else 0,
        engine="openpyxl",
    )

    loaded_df.columns = [
        str(column).strip()
        for column in loaded_df.columns
    ]

    missing_columns = [
        column
        for column in REQUIRED_COLUMNS
        if column not in loaded_df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing columns in Excel: "
            + ", ".join(missing_columns)
        )

    loaded_df = loaded_df[REQUIRED_COLUMNS].copy()

    loaded_df["Date"] = pd.to_datetime(
        loaded_df["Date"],
        errors="coerce",
    )

    loaded_df["Due Date"] = pd.to_datetime(
        loaded_df["Due Date"],
        errors="coerce",
    )

    text_columns = [
        "Task/ActionItem",
        "Responsible",
        "Area/Locations",
        "Department",
        "Priority",
        "Status",
    ]

    for column in text_columns:
        loaded_df[column] = (
            loaded_df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # Stable identifier used when updating filtered rows.
    loaded_df["_row_id"] = loaded_df.index

    return loaded_df


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(dataframe):
    save_df = dataframe.copy()

    save_df = save_df.drop(
        columns=["_row_id"],
        errors="ignore",
    )

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

        # Replace the source only after the temporary workbook is written.
        temporary_file.replace(EXCEL_FILE)
        load_excel.clear()

    except Exception:
        if temporary_file.exists():
            temporary_file.unlink()
        raise


# ============================================================
# HELPERS
# ============================================================

def format_date(value):
    if pd.isna(value):
        return ""

    return pd.Timestamp(value).strftime("%d-%b-%Y")


def unique_non_empty_values(series):
    return sorted(
        {
            str(value).strip()
            for value in series.dropna()
            if str(value).strip()
        }
    )


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="header">
        <h1>📊 Task Progress Dashboard</h1>
        <p>Excel-backed task management and progress monitoring</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# CHECK EXCEL
# ============================================================

if not EXCEL_FILE.exists():
    st.error(f"Excel file '{EXCEL_FILE}' was not found.")
    st.info(
        "Place tasks.xlsx in the same folder as app.py. "
        "The workbook must contain a sheet named 'Tasks'."
    )
    st.stop()


# ============================================================
# READ DATA
# ============================================================

try:
    df = load_excel(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime,
    )
except Exception as error:
    st.error(f"Unable to read Excel file: {error}")
    st.stop()


# ============================================================
# ADD NEW TASK
# ============================================================

with st.expander("➕ Add New Task", expanded=False):
    existing_departments = unique_non_empty_values(
        df["Department"]
    )

    with st.form(
        "add_task_form",
        clear_on_submit=True,
    ):
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
            new_task = pd.DataFrame(
                [
                    {
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
                    }
                ]
            )

            updated_df = pd.concat(
                [
                    df.drop(columns=["_row_id"], errors="ignore"),
                    new_task,
                ],
                ignore_index=True,
            )

            try:
                save_excel(updated_df)
                st.success(
                    f"✅ Task '{clean_task_name}' added successfully."
                )
                st.rerun()

            except PermissionError:
                st.error(
                    "Unable to update tasks.xlsx. Close the workbook "
                    "in Microsoft Excel and try again."
                )

            except Exception as error:
                st.error(f"Unable to add the task: {error}")


# ============================================================
# CALCULATE PROGRESS
# ============================================================

total_tasks = len(df)

completed_tasks = int(
    df["Status"]
    .str.casefold()
    .eq("close")
    .sum()
)

open_tasks = total_tasks - completed_tasks

completion_percentage = (
    (completed_tasks / total_tasks) * 100
    if total_tasks > 0
    else 0
)


# ============================================================
# TOP SECTION
# ============================================================

left_column, right_column = st.columns(
    [1.25, 1],
    gap="large",
)


# ============================================================
# LEFT - PROGRESS
# ============================================================

with left_column:
    st.markdown("### 📈 Overall Completion")

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
        unsafe_allow_html=True,
    )

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Total Tasks", total_tasks)
    metric2.metric("Open Tasks", open_tasks)
    metric3.metric("Completed", completed_tasks)


# ============================================================
# RIGHT - DEPARTMENT SUMMARY
# ============================================================

with right_column:
    st.markdown("### 🏢 Department-wise Tasks")

    if df.empty:
        st.info("No tasks are available yet.")
    else:
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
                Completed=("Completed", "sum"),
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
                    format="%.1f%%",
                )
            },
        )


# ============================================================
# COMPLETED TASKS
# ============================================================

st.divider()

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

if st.session_state.show_completed:
    st.markdown("### ✅ Completed Tasks")

    completed_df = df[
        df["Status"]
        .str.casefold()
        .eq("close")
    ].copy()

    if completed_df.empty:
        st.info("There are no completed tasks yet.")
    else:
        completed_df["Date"] = completed_df["Date"].apply(
            format_date
        )

        completed_df["Due Date"] = completed_df["Due Date"].apply(
            format_date
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
            hide_index=True,
        )

    st.divider()


# ============================================================
# OPEN TASKS
# ============================================================

st.markdown("### 📋 Open Tasks")

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
    expanded=True,
):
    filter1, filter2, filter3, filter4 = st.columns(4)

    search_text = filter1.text_input(
        "Search",
        placeholder="Task, responsible, location...",
    )

    department_options = unique_non_empty_values(
        open_df["Department"]
    )

    priority_options = unique_non_empty_values(
        open_df["Priority"]
    )

    status_options = unique_non_empty_values(
        open_df["Status"]
    )

    selected_departments = filter2.multiselect(
        "Department",
        department_options,
    )

    selected_priorities = filter3.multiselect(
        "Priority",
        priority_options,
    )

    selected_statuses = filter4.multiselect(
        "Status",
        status_options,
    )


# ============================================================
# APPLY FILTERS
# ============================================================

if search_text:
    search_value = search_text.casefold()

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
        .str.contains(search_value, regex=False)
    )

    open_df = open_df[search_mask]

if selected_departments:
    open_df = open_df[
        open_df["Department"].isin(selected_departments)
    ]

if selected_priorities:
    open_df = open_df[
        open_df["Priority"].isin(selected_priorities)
    ]

if selected_statuses:
    open_df = open_df[
        open_df["Status"].isin(selected_statuses)
    ]

st.caption(f"Showing {len(open_df)} open task(s).")


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
        errors="coerce",
    ).dt.date

    editable_data = table_data.drop(columns=["_row_id"])

    editor_priority_options = sorted(
        set(priority_options + PRIORITY_VALUES)
    )

    editor_status_options = sorted(
        set(status_options + STATUS_VALUES)
    )

    edited_data = st.data_editor(
        editable_data,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="task_editor",
        column_config={
            "Task/ActionItem": st.column_config.TextColumn(
                "Task",
                required=True,
            ),
            "Responsible": st.column_config.TextColumn(
                "Responsible"
            ),
            "Area/Locations": st.column_config.TextColumn(
                "Location"
            ),
            "Department": st.column_config.TextColumn(
                "Department"
            ),
            "Priority": st.column_config.SelectboxColumn(
                "Priority",
                options=editor_priority_options,
            ),
            "Due Date": st.column_config.DateColumn(
                "Due Date",
                format="DD-MM-YYYY",
            ),
            "Status": st.column_config.SelectboxColumn(
                "Status",
                options=editor_status_options,
            ),
        },
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
            updated_df = df.copy()

            for position in range(len(edited_data)):
                original_row_id = int(
                    table_data.iloc[position]["_row_id"]
                )

                row = edited_data.iloc[position]

                updated_df.loc[
                    original_row_id,
                    "Task/ActionItem",
                ] = str(row["Task/ActionItem"]).strip()

                updated_df.loc[
                    original_row_id,
                    "Responsible",
                ] = str(row["Responsible"]).strip()

                updated_df.loc[
                    original_row_id,
                    "Area/Locations",
                ] = str(row["Area/Locations"]).strip()

                updated_df.loc[
                    original_row_id,
                    "Department",
                ] = str(row["Department"]).strip()

                updated_df.loc[
                    original_row_id,
                    "Priority",
                ] = str(row["Priority"]).strip()

                updated_df.loc[
                    original_row_id,
                    "Due Date",
                ] = pd.to_datetime(
                    row["Due Date"],
                    errors="coerce",
                )

                updated_df.loc[
                    original_row_id,
                    "Status",
                ] = str(row["Status"]).strip()

            try:
                save_excel(updated_df)
                st.success("✅ Excel updated successfully.")
                st.rerun()

            except PermissionError:
                st.error(
                    "Unable to update tasks.xlsx. Close the workbook "
                    "in Microsoft Excel and try again."
                )

            except Exception as error:
                st.error(f"Unable to update Excel: {error}")

else:
    st.success("🎉 No open tasks match the current filters.")


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "A task is considered completed only when Status = Close."
)

st.caption(
    "Last dashboard refresh: "
    + datetime.now().strftime("%d-%b-%Y %H:%M:%S")
)
