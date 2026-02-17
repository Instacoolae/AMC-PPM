"""
PPM Productivity Data Collection Quick App
========================================

This Streamlit application provides a simple form for field technicians to
record their daily productivity for preventative maintenance (PPM) projects.

Key features:

* **Dynamic project selection** – technicians first choose the project owner.
  The project name drop‑down is filtered to only include projects owned by
  the selected owner. The emirate (city) associated with the project is
  displayed automatically.
* **Productivity inputs** – technicians enter how many indoor units, VRF
  outdoor units, DX outdoor units and AHU units they completed that day.
* **Team member selection** – technicians can choose up to three technicians
  and up to three helpers from the company technician list. A multi‑select
  widget is used to make it easy to pick multiple names.
* **Persistent storage** – each submission is appended to the “Inputs” sheet
  in the provided Excel workbook (`PPM App Data.xlsx`). This allows
  management to monitor field productivity without manual consolidation.

To run the app locally, install the required dependencies and execute

```bash
pip install streamlit openpyxl pandas
streamlit run ppm_quick_app.py
```

The app expects the file `PPM App Data.xlsx` to reside in the same directory.

"""

import datetime
from pathlib import Path

import pandas as pd
# Streamlit is only imported inside the main() function to avoid import
# errors when the module is imported for testing or in environments where
# streamlit is not installed. This design makes the utility functions
# (`load_data` and `save_submission`) usable without requiring streamlit.



def load_data(file_path: Path):
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
    # Try loading existing inputs; if it doesn't exist or is empty, create
    try:
        inputs_df = pd.read_excel(xls, sheet_name="Inputs")
        # Remove any legacy columns ending with 'Compelet' (misspelt count columns)
        cols_to_drop = [c for c in inputs_df.columns if c.endswith("Compelet")]
        if cols_to_drop:
            inputs_df = inputs_df.drop(columns=cols_to_drop)
    except ValueError:
        # Sheet does not exist – create an empty DataFrame with expected columns
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
    return project_df, technician_df, inputs_df


def save_submission(file_path: Path, inputs_df: pd.DataFrame):
    """Write the updated inputs DataFrame back to the Excel file.

    The function uses openpyxl via pandas ExcelWriter to either create the
    `Inputs` sheet or overwrite it while preserving other sheets.

    Parameters
    ----------
    file_path : Path
        Path to the Excel workbook.
    inputs_df : pd.DataFrame
        DataFrame containing all submissions.
    """
    # Write the cleaned inputs_df back to the Excel file. Use 'replace' mode for the
    # Inputs sheet so that any legacy columns are removed entirely rather than
    # overlaid. If the version of pandas does not support 'replace', it will
    # fall back to overlay which still writes all columns in inputs_df.
    with pd.ExcelWriter(
        file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        inputs_df.to_excel(writer, sheet_name="Inputs", index=False)


def main():
    """
    Run the PPM Productivity web app.

    This function loads the Excel workbook, displays a dynamic form for data
    collection and appends submissions to the workbook. It avoids using
    Streamlit forms so that dependent dropdowns update immediately when the
    project owner changes.
    """
    # Import streamlit inside the main function to avoid import errors when
    # running the module in environments without streamlit.
    import streamlit as st

    st.set_page_config(page_title="PPM Productivity App", layout="centered")
    st.title("PPM Productivity Data Collection")

    file_path = Path(__file__).parent / "PPM App Data.xlsx"
    if not file_path.exists():
        st.error(
            f"Could not find {file_path.name} in the current directory. Please place"
            " the Excel workbook in the same folder as this script."
        )
        st.stop()

    project_df, technician_df, inputs_df = load_data(file_path)
    # Clean up NaNs and ensure lists are unique
    project_df = project_df.dropna(subset=["Project Owner", "Project Name"])
    technician_df = technician_df.dropna(subset=["Technician Name"])

    owners = sorted(project_df["Project Owner"].dropna().unique())
    # Step 1: choose project owner (updates immediately)
    project_owner = st.selectbox(
        "Project Owner",
        owners,
        help="Choose the project owner. This filters the available projects.",
    )

    # Step 2: filter projects by selected owner
    project_options = project_df.loc[
        project_df["Project Owner"] == project_owner, "Project Name"
    ].dropna().unique()
    project_name = st.selectbox(
        "Project Name",
        project_options,
        help="Select the project name. The emirate will be looked up automatically.",
    )

    # Step 3: look up emirate for display
    emirate = ""
    if project_name:
        match = project_df.loc[
            (project_df["Project Owner"] == project_owner)
            & (project_df["Project Name"] == project_name)
        ]
        if not match.empty:
            emirate = match.iloc[0]["Emirate"]
    st.text_input("Emirate", value=emirate, disabled=True)

    # Step 4: numeric inputs for completed counts
    # Determine how many units remain for the selected project. This prevents
    # technicians from logging more units than exist in the project list.
    remaining_indoor = remaining_vrf = remaining_dx = remaining_ahu = 0
    if project_name:
        # Look up total quantities from the project list
        project_row = project_df.loc[
            (project_df["Project Owner"] == project_owner)
            & (project_df["Project Name"] == project_name)
        ]
        if not project_row.empty:
            total_indoor = int(project_row.iloc[0]["Indoors Qty"] or 0)
            total_vrf = int(project_row.iloc[0]["VRF OD Qty"] or 0)
            total_dx = int(project_row.iloc[0]["DX Outdoor Qty"] or 0)
            total_ahu = int(project_row.iloc[0]["AHU Qty"] or 0)
            # Sum completed so far from existing inputs
            inputs_for_project = inputs_df.loc[
                (inputs_df["Project Owner"] == project_owner)
                & (inputs_df["Project Name"] == project_name)
            ]
            completed_indoor = int(inputs_for_project["Indoors Completed"].sum())
            completed_vrf = int(inputs_for_project["VRF OD Completed"].sum())
            completed_dx = int(inputs_for_project["DX Outdoor Completed"].sum())
            completed_ahu = int(inputs_for_project["AHU Completed"].sum())
            # Compute remaining units (cannot go below zero)
            remaining_indoor = max(total_indoor - completed_indoor, 0)
            remaining_vrf = max(total_vrf - completed_vrf, 0)
            remaining_dx = max(total_dx - completed_dx, 0)
            remaining_ahu = max(total_ahu - completed_ahu, 0)

    # Render number inputs with dynamic maximums based on remaining counts
    indoors_completed = st.number_input(
        "Indoors Completed",
        min_value=0,
        max_value=remaining_indoor,
        step=1,
        value=0,
        help=f"Remaining indoor units: {remaining_indoor}"
    )
    vrf_completed = st.number_input(
        "VRF OD Completed",
        min_value=0,
        max_value=remaining_vrf,
        step=1,
        value=0,
        help=f"Remaining VRF outdoor units: {remaining_vrf}"
    )
    dx_completed = st.number_input(
        "DX Outdoor Completed",
        min_value=0,
        max_value=remaining_dx,
        step=1,
        value=0,
        help=f"Remaining DX outdoor units: {remaining_dx}"
    )
    ahu_completed = st.number_input(
        "AHU Completed",
        min_value=0,
        max_value=remaining_ahu,
        step=1,
        value=0,
        help=f"Remaining AHU units: {remaining_ahu}"
    )

    # Step 5: multi-select for technicians and helpers
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

    # Step 6: button to submit
    if st.button("Submit"):
        # Create a new row dictionary
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
        # Append to DataFrame
        inputs_df = pd.concat([inputs_df, pd.DataFrame([row])], ignore_index=True)
        # Persist to Excel
        save_submission(file_path, inputs_df)
        st.success("Data submitted successfully! Your entry has been saved.")

    # Display existing submissions for reference
    if not inputs_df.empty:
        st.subheader("Existing Submissions")
        # Drop any rows that are completely empty and remove duplicate entries
        display_df = inputs_df.dropna(how="all").drop_duplicates()
        # Show only the last 10 submissions for brevity
        st.dataframe(display_df.tail(10).astype(str), use_container_width=True)

        # Offer a download button for the cleaned data
        csv_data = display_df.to_csv(index=False)
        st.download_button(
            label="Download CSV",
            data=csv_data,
            file_name="PPM_inputs_export.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()