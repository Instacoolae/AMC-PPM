"""
PPM Productivity Data Collection Application – Simplified Login
----------------------------------------------------------------

This version of the PPM productivity app removes OTP verification from the
login/registration process. Users authenticate themselves by entering
their mobile number. If the number exists in the `users.csv` registry
and has a role of `admin`, the user receives administrator privileges
(access to the Summary & Export tab). Otherwise, the user is treated
as a regular technician. Unknown numbers are automatically added to
`users.csv` with a role of `user`.

Key features:

* **User registration without OTP** – simply enter your phone number
  to sign in. The app looks up the number in the users registry.
  If not found, it creates a new user entry with role "user".
* **Role‑based permissions** – administrators can view progress
  summaries and download cleaned inputs. Regular users can only
  submit entries.
* **Clean data model and dynamic quotas** – identical to previous
  versions. Legacy “Compelet” columns are removed, and the number
  inputs enforce remaining quantities per project.
* **Instacool branding** – light blue colour palette matching the
  reference site.

Note: Without verification, anyone knowing an admin's phone number can
gain administrator access. Use this variant only when you trust
participants or while experimenting. For production use, integrate
proper verification via SMS or other methods.
"""

import datetime
from pathlib import Path
from typing import Tuple

import pandas as pd


# Paths to the Excel workbook and user registry. The workbook must
# reside in the same folder as this script. The users file is a CSV
# maintained by the application to store registered users and their
# roles. Administrators have the role "admin". Normal users have
# role "user".
DATA_FILE = Path(__file__).parent / "PPM App Data.xlsx"
USERS_FILE = Path(__file__).parent / "users.csv"


def load_users() -> pd.DataFrame:
    """Load registered users from USERS_FILE.

    Returns
    -------
    pd.DataFrame
        DataFrame with columns ["phone", "name", "role"]. If the file
        does not exist, returns an empty DataFrame with these columns.
    """
    if USERS_FILE.exists():
        try:
            users = pd.read_csv(USERS_FILE, dtype=str)
        except Exception:
            users = pd.DataFrame(columns=["phone", "name", "role"])
    else:
        users = pd.DataFrame(columns=["phone", "name", "role"])
    return users


def save_users(users: pd.DataFrame) -> None:
    """Persist the users DataFrame to USERS_FILE."""
    users.to_csv(USERS_FILE, index=False)


def load_data(file_path: Path) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Load project list, technician list and inputs from the Excel file.

    Parameters
    ----------
    file_path : Path
        Path to the Excel workbook containing the sheets.

    Returns
    -------
    tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]
        DataFrames for project list, technician list and inputs sheet.
    """
    xls = pd.ExcelFile(file_path)
    project_df = pd.read_excel(xls, sheet_name="Project List")
    technician_df = pd.read_excel(xls, sheet_name="Technician List")
    try:
        inputs_df = pd.read_excel(xls, sheet_name="Inputs")
        # Drop misspelt "Compelet" columns if present
        cols_to_drop = [c for c in inputs_df.columns if c.endswith("Compelet")]
        if cols_to_drop:
            inputs_df = inputs_df.drop(columns=cols_to_drop)
    except Exception:
        # Create an empty DataFrame with expected columns
        inputs_df = pd.DataFrame(
            columns=[
                "Date",
                "Project Owner",
                "Project Name",
                "Emirate",
                "Indoors Completed",
                "VRF OD Completed",
                "DX Outdoor Completed",
                "AHU Completed",
                "Technician name 1",
                "Technician name 2",
                "Technician name 3",
                "Helper name 1",
                "Helper name 2",
                "Helper name 3",
            ]
        )
    # Clean NaNs and duplicates in project and technician lists
    project_df = project_df.dropna(subset=["Project Owner", "Project Name"])
    technician_df = technician_df.dropna(subset=["Technician Name"])
    return project_df, technician_df, inputs_df


def save_inputs(inputs_df: pd.DataFrame) -> None:
    """Write the updated inputs DataFrame back to the Excel file.

    This replaces the existing "Inputs" sheet entirely, ensuring legacy
    columns are removed. Other sheets in the workbook are preserved.
    """
    with pd.ExcelWriter(
        DATA_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        inputs_df.to_excel(writer, sheet_name="Inputs", index=False)


def app_header():
    """Inject custom CSS for branding and set up the page title."""
    import streamlit as st

    # Define Instacool‑inspired colours
    primary = "#00ADEF"  # bright blue similar to Instacool
    secondary = "#F2F9FC"  # very light blue background
    accent = "#0090C6"  # darker blue for buttons/hover

    st.set_page_config(page_title="PPM Productivity App", layout="wide")
    # Custom styles injected via HTML
    st.markdown(
        f"""
        <style>
        .reportview-container .main {{
            background-color: {secondary};
        }}
        /* Primary colour for buttons */
        .stButton>button {{
            background-color: {primary};
            color: white;
            border-radius: 4px;
        }}
        .stButton>button:hover {{
            background-color: {accent};
            color: white;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("PPM Productivity Data Collection")


def show_login(users: pd.DataFrame) -> Tuple[bool, bool, str]:
    """Render a simplified login/registration page without OTP.

    Parameters
    ----------
    users : pd.DataFrame
        DataFrame containing registered users.

    Returns
    -------
    tuple[bool, bool, str]
        Tuple `(authenticated, is_admin, phone)` where `authenticated` is
        True when the user has entered their phone number and the app
        has determined their role, `is_admin` indicates admin privilege
        based on the user record, and `phone` is the logged in phone
        number.
    """
    import streamlit as st

    # Initialize session state keys
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.is_admin = False
        st.session_state.phone = ''

    authenticated = False
    is_admin = False
    phone = ''

    if not st.session_state.logged_in:
        st.subheader("Login / Register")
        phone_input = st.text_input("Enter your mobile number", value=st.session_state.get('phone', ''))
        if st.button("Login"):
            phone = phone_input.strip()
            if phone:
                # Look up the user
                if phone in users['phone'].values:
                    user_record = users[users['phone'] == phone].iloc[0]
                    is_admin = user_record['role'].lower() == 'admin'
                else:
                    # Register new user
                    new_row = pd.DataFrame([{'phone': phone, 'name': '', 'role': 'user'}])
                    users = pd.concat([users, new_row], ignore_index=True)
                    save_users(users)
                    is_admin = False
                authenticated = True
                st.session_state.logged_in = True
                st.session_state.phone = phone
                st.session_state.is_admin = is_admin
    else:
        authenticated = True
        phone = st.session_state.phone
        is_admin = st.session_state.is_admin

    return authenticated, is_admin, phone


def main() -> None:
    """Run the simplified Streamlit application."""
    import streamlit as st

    # Branding header
    app_header()

    # Load users registry
    users = load_users()

    # Handle authentication (without OTP)
    authenticated, is_admin, phone = show_login(users)
    if not authenticated:
        st.stop()

    # Load project and input data
    project_df, technician_df, inputs_df = load_data(DATA_FILE)

    # Prepare options
    owners = sorted(project_df["Project Owner"].dropna().unique())

    st.subheader(f"Welcome {'Admin' if is_admin else 'User'}")

    # Tabs for data entry and (for admin) reports
    tabs = ["Submit Entry"]
    if is_admin:
        tabs.append("Summary & Export")
    selected_tab = st.radio("Navigation", tabs)

    if selected_tab == "Submit Entry":
        # Data entry form
        project_owner = st.selectbox(
            "Project Owner",
            owners,
            help="Choose the project owner. This filters the available projects.",
        )
        # Filter projects by owner
        project_options = project_df.loc[
            project_df["Project Owner"] == project_owner, "Project Name"
        ].dropna().unique()
        project_name = st.selectbox(
            "Project Name",
            project_options,
            help="Select the project name. The emirate will be looked up automatically.",
        )
        # Look up emirate
        emirate = ""
        if project_name:
            match = project_df.loc[
                (project_df["Project Owner"] == project_owner)
                & (project_df["Project Name"] == project_name)
            ]
            if not match.empty:
                emirate = match.iloc[0]["Emirate"]
        st.text_input("Emirate", value=emirate, disabled=True)
        # Calculate remaining quotas
        remaining_indoor = remaining_vrf = remaining_dx = remaining_ahu = 0
        if project_name:
            total_row = project_df.loc[
                (project_df["Project Owner"] == project_owner)
                & (project_df["Project Name"] == project_name)
            ]
            if not total_row.empty:
                total_indoor = int(total_row.iloc[0]["Indoors Qty"] or 0)
                total_vrf = int(total_row.iloc[0]["VRF OD Qty"] or 0)
                total_dx = int(total_row.iloc[0]["DX Outdoor Qty"] or 0)
                total_ahu = int(total_row.iloc[0]["AHU Qty"] or 0)
                # Sum completed so far
                prev = inputs_df.loc[
                    (inputs_df["Project Owner"] == project_owner)
                    & (inputs_df["Project Name"] == project_name)
                ]
                completed_indoor = int(prev["Indoors Completed"].sum())
                completed_vrf = int(prev["VRF OD Completed"].sum())
                completed_dx = int(prev["DX Outdoor Completed"].sum())
                completed_ahu = int(prev["AHU Completed"].sum())
                remaining_indoor = max(total_indoor - completed_indoor, 0)
                remaining_vrf = max(total_vrf - completed_vrf, 0)
                remaining_dx = max(total_dx - completed_dx, 0)
                remaining_ahu = max(total_ahu - completed_ahu, 0)
        # Inputs for completed counts with restrictions
        st.write(
            f"Remaining Indoors: {remaining_indoor}, VRF: {remaining_vrf}, DX: {remaining_dx}, AHU: {remaining_ahu}"
        )
        indoors_completed = st.number_input(
            "Indoors Completed",
            min_value=0,
            max_value=remaining_indoor,
            step=1,
            value=0,
        )
        vrf_completed = st.number_input(
            "VRF OD Completed",
            min_value=0,
            max_value=remaining_vrf,
            step=1,
            value=0,
        )
        dx_completed = st.number_input(
            "DX Outdoor Completed",
            min_value=0,
            max_value=remaining_dx,
            step=1,
            value=0,
        )
        ahu_completed = st.number_input(
            "AHU Completed",
            min_value=0,
            max_value=remaining_ahu,
            step=1,
            value=0,
        )
        # Technician and helper selection
        technician_names = sorted(technician_df["Technician Name"].dropna().unique())
        tech_selected = st.multiselect(
            "Technician(s)",
            technician_names,
            max_selections=3,
            help="Select up to three technicians involved in this task.",
        )
        helper_selected = st.multiselect(
            "Helper(s)",
            technician_names,
            max_selections=3,
            help="Select up to three helpers assisting in this task.",
        )
        # Submission button
        if st.button("Submit"):
            # Construct new row
            row = {
                "Date": datetime.date.today(),
                "Project Owner": project_owner,
                "Project Name": project_name,
                "Emirate": emirate,
                "Indoors Completed": int(indoors_completed),
                "VRF OD Completed": int(vrf_completed),
                "DX Outdoor Completed": int(dx_completed),
                "AHU Completed": int(ahu_completed),
                "Technician name 1": tech_selected[0] if len(tech_selected) > 0 else "",
                "Technician name 2": tech_selected[1] if len(tech_selected) > 1 else "",
                "Technician name 3": tech_selected[2] if len(tech_selected) > 2 else "",
                "Helper name 1": helper_selected[0] if len(helper_selected) > 0 else "",
                "Helper name 2": helper_selected[1] if len(helper_selected) > 1 else "",
                "Helper name 3": helper_selected[2] if len(helper_selected) > 2 else "",
            }
            # Append to DataFrame and save
            inputs_df = pd.concat([inputs_df, pd.DataFrame([row])], ignore_index=True)
            save_inputs(inputs_df)
            st.success("Submission saved successfully!")

        # Display recent submissions with technician and helper names
        if not inputs_df.empty:
            st.subheader("Recent Submissions")
            # Show the last 10 entries with all columns converted to string to avoid
            # Arrow conversion issues in Streamlit's dataframe display.
            st.dataframe(inputs_df.tail(10).astype(str), use_container_width=True)

    elif selected_tab == "Summary & Export" and is_admin:
        # Admin view: show project progress and allow download
        st.subheader("Project Progress Summary")
        # For each project, compute completion percentage
        summary_rows = []
        for _, p_row in project_df.iterrows():
            owner = p_row["Project Owner"]
            proj = p_row["Project Name"]
            total_indoor = int(p_row.get("Indoors Qty", 0) or 0)
            total_vrf = int(p_row.get("VRF OD Qty", 0) or 0)
            total_dx = int(p_row.get("DX Outdoor Qty", 0) or 0)
            total_ahu = int(p_row.get("AHU Qty", 0) or 0)
            prev = inputs_df.loc[
                (inputs_df["Project Owner"] == owner) & (inputs_df["Project Name"] == proj)
            ]
            completed_indoor = int(prev["Indoors Completed"].sum())
            completed_vrf = int(prev["VRF OD Completed"].sum())
            completed_dx = int(prev["DX Outdoor Completed"].sum())
            completed_ahu = int(prev["AHU Completed"].sum())
            summary_rows.append({
                "Project Owner": owner,
                "Project Name": proj,
                "Indoors Completed": completed_indoor,
                "Indoors Total": total_indoor,
                "VRF OD Completed": completed_vrf,
                "VRF OD Total": total_vrf,
                "DX Outdoor Completed": completed_dx,
                "DX Outdoor Total": total_dx,
                "AHU Completed": completed_ahu,
                "AHU Total": total_ahu,
            })
        summary_df = pd.DataFrame(summary_rows)
        # Display summary table with completion percentages
        st.dataframe(summary_df, use_container_width=True)
        # Show progress bars for each project
        for _, row in summary_df.iterrows():
            st.markdown(f"### {row['Project Owner']} - {row['Project Name']}")
            for comp_key, total_key, label in [
                ("Indoors Completed", "Indoors Total", "Indoor Units"),
                ("VRF OD Completed", "VRF OD Total", "VRF Outdoor"),
                ("DX Outdoor Completed", "DX Outdoor Total", "DX Outdoor"),
                ("AHU Completed", "AHU Total", "AHU"),
            ]:
                total_val = row[total_key]
                comp_val = row[comp_key]
                if total_val > 0:
                    # Compute progress ratio and clamp it between 0 and 1
                    pct = comp_val / total_val if total_val else 0
                    # Ensure the progress bar receives a value within [0, 1]
                    pct_clamped = max(0.0, min(1.0, pct))
                    st.progress(pct_clamped)
                    st.write(f"{label}: {comp_val} / {total_val} ({pct_clamped:.0%})")
        # Export cleaned inputs data for admin
        cleaned = inputs_df.drop_duplicates()
        csv_data = cleaned.to_csv(index=False)
        st.download_button(
            label="Download CSV",
            data=csv_data,
            file_name="PPM_inputs_export.csv",
            mime="text/csv",
        )

        # Display full submissions table for admin including technicians and helpers
        if not inputs_df.empty:
            st.subheader("All Submissions")
            st.dataframe(inputs_df.astype(str), use_container_width=True)


if __name__ == "__main__":
    main()