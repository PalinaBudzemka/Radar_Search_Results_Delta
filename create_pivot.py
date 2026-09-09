import pandas as pd
import glob
import os
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime

if len(sys.argv) > 1:
    input_file = sys.argv[1]
else:
    input_files = glob.glob("search_results/*.xlsx")
    if not input_files:
        raise FileNotFoundError("No Excel files found in search_results")

    input_file = max(input_files, key=os.path.getmtime)

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
output_file = f"pivot_results/pivot_output_{timestamp}.xlsx"

# Read data from Positions sheet
df = pd.read_excel(input_file, sheet_name="Positions")

# Count positions by Customer and Project
project_counts = (
    df.groupby(["Customer", "Project"], dropna=False)["Position ID"]
    .count()
    .reset_index(name="Count of Position ID")
)

# Calculate total positions by Customer
customer_totals = (
    project_counts.groupby("Customer", dropna=False)["Count of Position ID"]
    .sum()
    .reset_index(name="Customer Total")
)

# Sort customers by total count DESC
customer_totals = customer_totals.sort_values(
    by=["Customer Total", "Customer"],
    ascending=[False, True]
)

# Build final summary rows
summary_rows = []

for _, customer_row in customer_totals.iterrows():
    customer = customer_row["Customer"]
    customer_total = customer_row["Customer Total"]

    # Add customer total row
    summary_rows.append({
        "Customer": customer,
        "Project": "",
        "Count of Position ID": customer_total,
        "Row Type": "Customer Total"
    })

    # Get projects for this customer, sorted by count DESC
    customer_projects = project_counts[project_counts["Customer"] == customer].sort_values(
        by=["Count of Position ID", "Project"],
        ascending=[False, True]
    )

    # Add project rows
    for _, project_row in customer_projects.iterrows():
        summary_rows.append({
            "Customer": project_row["Customer"],
            "Project": project_row["Project"],
            "Count of Position ID": project_row["Count of Position ID"],
            "Row Type": "Project"
        })

# Add Grand Total row
grand_total = project_counts["Count of Position ID"].sum()

summary_rows.append({
    "Customer": "Grand Total",
    "Project": "",
    "Count of Position ID": grand_total,
    "Row Type": "Grand Total"
})

summary_df = pd.DataFrame(summary_rows)

# Keep Row Type only for formatting, not for final visible output
visible_summary_df = summary_df[["Customer", "Project", "Count of Position ID"]]

# Save original data + summary
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Original Positions", index=False)
    visible_summary_df.to_excel(writer, sheet_name="Summary", index=False)

# Formatting
wb = load_workbook(output_file)
ws = wb["Summary"]

header_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
subtotal_fill = PatternFill(start_color="EAF4E2", end_color="EAF4E2", fill_type="solid")
grand_total_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")

# Header formatting
for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.fill = header_fill

# Format customer total rows and grand total row
for excel_row_number, row_type in enumerate(summary_df["Row Type"], start=2):
    if row_type == "Customer Total":
        for cell in ws[excel_row_number]:
            cell.font = Font(bold=True)
            cell.fill = subtotal_fill

    elif row_type == "Grand Total":
        for cell in ws[excel_row_number]:
            cell.font = Font(bold=True)
            cell.fill = grand_total_fill

# Auto-adjust column widths
for column_cells in ws.columns:
    max_length = 0
    column_letter = get_column_letter(column_cells[0].column)

    for cell in column_cells:
        if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))

    ws.column_dimensions[column_letter].width = max_length + 2

wb.save(output_file)

print(f"Created {output_file}")