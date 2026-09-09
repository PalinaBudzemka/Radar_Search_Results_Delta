import glob
import os
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# Folder where pivot_output_*.xlsx files are located
pivot_folder = "pivot_results"

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_file = f"delta_results/delta_results_{timestamp}.xlsx"


def find_latest_two_pivot_files(folder):
    files = glob.glob(os.path.join(folder, "pivot_output_*.xlsx"))

    if len(files) < 2:
        raise Exception("Need at least two pivot_output_*.xlsx files to compare.")

    # Because filenames contain timestamp, sorting by filename works correctly
    files = sorted(files, key=lambda x: os.path.basename(x), reverse=True)

    latest = files[0]
    previous = files[1]

    return latest, previous


def read_positions(file_path):
    df = pd.read_excel(file_path, sheet_name="Original Positions")

    if "Position ID" not in df.columns:
        raise Exception(f"'Position ID' column not found in {file_path}")

    df["Position ID Compare"] = df["Position ID"].astype(str).str.strip()

    # Remove rows without Position ID
    df = df[df["Position ID Compare"].notna()]
    df = df[df["Position ID Compare"] != ""]

    # Avoid duplicate Position IDs
    df = df.drop_duplicates(subset=["Position ID Compare"]).copy()

    return df


def compare_positions(latest_file, previous_file):
    latest_df = read_positions(latest_file)
    previous_df = read_positions(previous_file)

    latest_ids = set(latest_df["Position ID Compare"])
    previous_ids = set(previous_df["Position ID Compare"])

    added_ids = latest_ids - previous_ids
    removed_ids = previous_ids - latest_ids
    unchanged_ids = latest_ids & previous_ids

    added_df = latest_df[latest_df["Position ID Compare"].isin(added_ids)].copy()
    added_df["IsChanged"] = "Changed"
    added_df["Change"] = "Added"

    removed_df = previous_df[previous_df["Position ID Compare"].isin(removed_ids)].copy()
    removed_df["IsChanged"] = "Changed"
    removed_df["Change"] = "Removed"

    unchanged_df = latest_df[latest_df["Position ID Compare"].isin(unchanged_ids)].copy()
    unchanged_df["IsChanged"] = "Unchanged"
    unchanged_df["Change"] = ""

    delta_df = pd.concat(
        [added_df, removed_df, unchanged_df],
        ignore_index=True
    )

    delta_df = delta_df.drop(columns=["Position ID Compare"])

    sort_columns = []
    for column in ["IsChanged", "Change", "Customer", "Project", "Position ID"]:
        if column in delta_df.columns:
            sort_columns.append(column)

    if sort_columns:
        delta_df = delta_df.sort_values(by=sort_columns)

    return delta_df, len(added_ids), len(removed_ids), len(unchanged_ids), len(latest_ids), len(previous_ids)


def read_summary(file_path):
    df = pd.read_excel(file_path, sheet_name="Summary")

    required_columns = ["Customer", "Project", "Count of Position ID"]

    for column in required_columns:
        if column not in df.columns:
            raise Exception(f"'{column}' column not found in Summary sheet of {file_path}")

    df["Customer"] = df["Customer"].fillna("").astype(str).str.strip()
    df["Project"] = df["Project"].fillna("").astype(str).str.strip()
    df["Count of Position ID"] = pd.to_numeric(
        df["Count of Position ID"],
        errors="coerce"
    ).fillna(0).astype(int)

    # Remove Grand Total row
    df = df[df["Customer"] != "Grand Total"].copy()

    return df


def compare_summary_tabs(latest_file, previous_file):
    latest_summary = read_summary(latest_file)
    previous_summary = read_summary(previous_file)

    # ----------------------------
    # Compare Customer totals
    # ----------------------------

    # Customer total rows are rows where Project is empty
    latest_customers = latest_summary[latest_summary["Project"] == ""].copy()
    previous_customers = previous_summary[previous_summary["Project"] == ""].copy()

    latest_customers = latest_customers[["Customer", "Count of Position ID"]].rename(
        columns={"Count of Position ID": "Latest Count"}
    )

    previous_customers = previous_customers[["Customer", "Count of Position ID"]].rename(
        columns={"Count of Position ID": "Previous Count"}
    )

    customer_compare = previous_customers.merge(
        latest_customers,
        on="Customer",
        how="outer"
    )

    customer_compare["Previous Count"] = customer_compare["Previous Count"].fillna(0).astype(int)
    customer_compare["Latest Count"] = customer_compare["Latest Count"].fillna(0).astype(int)
    customer_compare["Delta"] = customer_compare["Latest Count"] - customer_compare["Previous Count"]

    added_customers = customer_compare[
        (customer_compare["Previous Count"] == 0) &
        (customer_compare["Latest Count"] > 0)
    ].copy()

    added_customers["Change"] = "Added"

    removed_customers = customer_compare[
        (customer_compare["Previous Count"] > 0) &
        (customer_compare["Latest Count"] == 0)
    ].copy()

    removed_customers["Change"] = "Removed"

    added_customers = added_customers[
        ["Customer", "Previous Count", "Latest Count", "Delta", "Change"]
    ].sort_values(by=["Latest Count", "Customer"], ascending=[False, True])

    removed_customers = removed_customers[
        ["Customer", "Previous Count", "Latest Count", "Delta", "Change"]
    ].sort_values(by=["Previous Count", "Customer"], ascending=[False, True])

    # Customers existing in both latest and previous files
    existing_customers = customer_compare[
        (customer_compare["Previous Count"] > 0) &
        (customer_compare["Latest Count"] > 0)
    ]["Customer"].tolist()

    # ----------------------------
    # Compare Customer + Project counts
    # ----------------------------

    # Project rows are rows where Project is not empty
    latest_projects = latest_summary[latest_summary["Project"] != ""].copy()
    previous_projects = previous_summary[previous_summary["Project"] != ""].copy()

    latest_projects = latest_projects[
        ["Customer", "Project", "Count of Position ID"]
    ].rename(columns={"Count of Position ID": "Latest Count"})

    previous_projects = previous_projects[
        ["Customer", "Project", "Count of Position ID"]
    ].rename(columns={"Count of Position ID": "Previous Count"})

    project_compare = previous_projects.merge(
        latest_projects,
        on=["Customer", "Project"],
        how="outer"
    )

    project_compare["Previous Count"] = project_compare["Previous Count"].fillna(0).astype(int)
    project_compare["Latest Count"] = project_compare["Latest Count"].fillna(0).astype(int)
    project_compare["Delta"] = project_compare["Latest Count"] - project_compare["Previous Count"]

    # Important:
    # Exclude projects belonging to fully added or fully removed customers.
    # Keep only customers that existed in both files.
    project_compare = project_compare[
        project_compare["Customer"].isin(existing_customers)
    ].copy()

    decreased_projects = project_compare[project_compare["Delta"] < 0].copy()
    decreased_projects["Change"] = "Decreased"

    increased_projects = project_compare[project_compare["Delta"] > 0].copy()
    increased_projects["Change"] = "Increased"

    decreased_projects = decreased_projects[
        ["Customer", "Project", "Previous Count", "Latest Count", "Delta", "Change"]
    ].sort_values(by=["Delta", "Customer", "Project"], ascending=[True, True, True])

    increased_projects = increased_projects[
        ["Customer", "Project", "Previous Count", "Latest Count", "Delta", "Change"]
    ].sort_values(by=["Delta", "Customer", "Project"], ascending=[False, True, True])

    return added_customers, removed_customers, decreased_projects, increased_projects


def format_workbook(file_path):
    wb = load_workbook(file_path)

    header_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    added_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    removed_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    increased_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    decreased_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")

    for ws in wb.worksheets:
        # Header formatting
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill

        headers = [cell.value for cell in ws[1]]

        # Conditional row coloring
        if "Change" in headers:
            change_col_idx = headers.index("Change") + 1

            for row_number in range(2, ws.max_row + 1):
                change_value = ws.cell(row=row_number, column=change_col_idx).value

                if change_value == "Added":
                    fill = added_fill
                elif change_value == "Removed":
                    fill = removed_fill
                elif change_value == "Increased":
                    fill = increased_fill
                elif change_value == "Decreased":
                    fill = decreased_fill
                else:
                    fill = None

                if fill:
                    for cell in ws[row_number]:
                        cell.fill = fill

        # Auto-adjust column widths
        for column_cells in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column_cells[0].column)

            for cell in column_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))

            ws.column_dimensions[column_letter].width = max_length + 2

    wb.save(file_path)


latest_file, previous_file = find_latest_two_pivot_files(pivot_folder)

print(f"Latest file:   {latest_file}")
print(f"Previous file: {previous_file}")

delta_positions_df, added_count, removed_count, unchanged_count, latest_total, previous_total = compare_positions(
    latest_file,
    previous_file
)

added_customers_df, removed_customers_df, decreased_projects_df, increased_projects_df = compare_summary_tabs(
    latest_file,
    previous_file
)

run_summary_df = pd.DataFrame({
    "Metric": [
        "Latest file",
        "Previous file",
        "Added positions",
        "Removed positions",
        "Unchanged positions",
        "Total positions in latest",
        "Total positions in previous",
        "Added customers",
        "Removed customers",
        "Customer/Project increases",
        "Customer/Project decreases"
    ],
    "Value": [
        os.path.basename(latest_file),
        os.path.basename(previous_file),
        added_count,
        removed_count,
        unchanged_count,
        latest_total,
        previous_total,
        len(added_customers_df),
        len(removed_customers_df),
        len(increased_projects_df),
        len(decreased_projects_df)
    ]
})

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    run_summary_df.to_excel(writer, sheet_name="Run Summary", index=False)
    delta_positions_df.to_excel(writer, sheet_name="Delta Positions", index=False)
    added_customers_df.to_excel(writer, sheet_name="Added Customers", index=False)
    removed_customers_df.to_excel(writer, sheet_name="Removed Customers", index=False)
    decreased_projects_df.to_excel(writer, sheet_name="Decreased Counts", index=False)
    increased_projects_df.to_excel(writer, sheet_name="Increased Counts", index=False)

format_workbook(output_file)

print(f"Created {output_file}")