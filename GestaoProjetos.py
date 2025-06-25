
# -*- main.py -*-
# -------------------------------------------
# Project Management    - Version 1.0
# Author: Ermelino Piazzetta
# Creation Date: 2025-06-25
# Description: Main entry point of the system.
# -------------------------------------------

from datetime import datetime
from input_utils import prompt_for_value, register_items
from file_utils import open_file, list_existing_projects
from spreadsheet import save_project_spreadsheet, save_project_info_sheet
from summary import update_project_summary

def collect_project_info() -> dict:
    print("\n=== New Project Information ===")
    manager = input("Project manager name: ").strip()
    while True:
        opening_date_str = input("Opening date (YYYY-MM-DD): ").strip()
        try:
            datetime.strptime(opening_date_str, "%Y-%m-%d")
            break
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD.")
    est_completion = input("Estimated completion date (optional, YYYY-MM-DD): ").strip()
    cost_estimate = prompt_for_value("Estimated cost (R$): ", float)

    return {
        "Manager": manager,
        "Opening Date": opening_date_str,
        "Estimated Completion": est_completion if est_completion else "N/A",
        "Estimated Cost": f"R$ {cost_estimate:,.2f}"
    }

def main() -> None:
    print("=== Project Registration System ===")

    project_info = None

    while True:
        choice = input("\nStart a (N)ew project or open an (E)xisting one? (n/e): ").strip().lower()
        if choice == 'n':
            project_name = input("Enter the new project name: ").strip().replace(" ", "_")
            project_info = collect_project_info()
            break
        elif choice == 'e':
            projects = list_existing_projects()
            if not projects:
                print("No existing projects found.")
                return
            print("Existing projects:")
            for i, project in enumerate(projects, start=1):
                print(f"{i}. {project}")
            index = prompt_for_value("Select a project number: ", int)
            if 1 <= index <= len(projects):
                project_name = projects[index - 1]
                break
            else:
                print("Invalid choice.")
        else:
            print("Invalid option. Enter 'n' or 'e'.")

    while True:
        items, totals = register_items()
        if items:
            total = save_project_spreadsheet(project_name, items, totals)
            if project_info:
                save_project_info_sheet(project_name, project_info)
            update_project_summary(project_name, total)
        else:
            print("No items registered.")

        cont = input("\nRegister more items for this project? (y/n): ").strip().lower()
        if cont != 'y':
            print("Ending project registration.")
            break

if __name__ == "__main__":
    main()


# -*- input_utils.py -*-
# -------------------------------------------
# Project Management    - Version 1.0
# Author: Ermelino Piazzetta
# Creation Date: 2025-06-25
# Description: Input and validation utilities.
# -------------------------------------------

def prompt_for_value(message: str, val_type: type):
    while True:
        user_input = input(message)
        try:
            return val_type(user_input)
        except ValueError:
            print(f"Invalid input. Please enter a value of type {val_type.__name__}.")

def register_items():
    print("\n=== Item Registration ===")
    print("Enter 'end' in the description to finish.\n")

    items = []
    totals_by_category = {}

    while True:
        description = input("Item description (or 'end' to finish): ").strip()
        if description.lower() == 'end':
            break

        quantity = prompt_for_value("Quantity: ", float)
        unit = input("Unit (e.g. man-hour, material): ").strip()
        unit_price = prompt_for_value("Unit price (R$): ", float)
        total_price = quantity * unit_price

        items.append({
            "Description": description,
            "Quantity": quantity,
            "Unit": unit,
            "Unit Price": unit_price,
            "Total Price": total_price
        })

        totals_by_category[description] = totals_by_category.get(description, 0) + total_price

        print("Item successfully added!\n")

    return items, totals_by_category

# -*- file_utils.py -*-
# -------------------------------------------
# Project Management    - Version 1.0
# Author: Ermelino Piazzetta
# Creation Date: 2025-06-25
# Description: File handling utilities (open, list).
# -------------------------------------------

import os
import platform
from typing import List

def open_file(file_name: str) -> None:
    system = platform.system()
    if system == "Windows":
        os.startfile(file_name)
    elif system == "Darwin":
        os.system(f"open '{file_name}'")
    else:
        os.system(f"xdg-open '{file_name}'")

def list_existing_projects() -> List[str]:
    return [f[8:-5] for f in os.listdir() if f.startswith("projeto_") and f.endswith(".xlsx")]

# -*- spreadsheet.py -*-
# -------------------------------------------
# Project Management    - Version 1.0
# Author: Ermelino Piazzetta
# Creation Date: 2025-06-25
# Description: Handles spreadsheet creation, formatting, charts.
# -------------------------------------------

import os
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.chart import BarChart, Reference

def apply_formatting(ws):
    bold_font = Font(bold=True)
    fill_header = PatternFill("solid", fgColor="BDD7EE")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    money_style = NamedStyle(name="money_style", number_format='"R$"#,##0.00')
    if "money_style" not in ws.parent.named_styles:
        ws.parent.add_named_style(money_style)

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=5):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.row == 1:
                cell.font = bold_font
                cell.fill = fill_header
            elif cell.column in (4, 5):
                cell.style = money_style

def add_chart(ws, start_row: int):
    chart = BarChart()
    chart.title = "Totals by Category"
    chart.x_axis.title = "Category"
    chart.y_axis.title = "Total Value (R$)"

    data = Reference(ws, min_col=2, min_row=start_row + 1, max_row=ws.max_row)
    categories = Reference(ws, min_col=1, min_row=start_row + 2, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    ws.add_chart(chart, f"A{ws.max_row + 3}")

def write_items_to_sheet(ws, items, totals_by_category):
    for item in items:
        ws.append([
            item["Description"],
            item["Quantity"],
            item["Unit"],
            item["Unit Price"],
            item["Total Price"]
        ])
    ws.append([])
    totals_start_row = ws.max_row + 1
    ws.append(["Totals by Category"])
    ws.append(["Category", "Total (R$)"])
    for category, total in totals_by_category.items():
        ws.append([category, total])

    apply_formatting(ws)
    add_chart(ws, totals_start_row)

def save_project_spreadsheet(project_name: str, items, totals_by_category) -> float:
    filename = f"projeto_{project_name}.xlsx"
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Project Items"
        ws.append(["Description", "Quantity", "Unit", "Unit Price (R$)", "Total Price (R$)"])

    write_items_to_sheet(ws, items, totals_by_category)
    wb.save(filename)
    print(f"\nFile '{filename}' saved successfully.")
    from file_utils import open_file
    open_file(filename)

    return sum(totals_by_category.values())

def save_project_info_sheet(project_name: str, info: dict) -> None:
    filename = f"projeto_{project_name}.xlsx"
    if os.path.exists(filename):
        wb = load_workbook(filename)
    else:
        wb = Workbook()

    # Remove 'Project Info' sheet if exists
    if "Project Info" in wb.sheetnames:
        std = wb["Project Info"]
        wb.remove(std)

    ws = wb.create_sheet(title="Project Info")

    ws.append(["Field", "Value"])
    for key, value in info.items():
        ws.append([key, value])

    wb.save(filename)
    print(f"Project info sheet updated in '{filename}'.")

# -*- summary.py -*-
# -------------------------------------------
# Project Management    - Version 1.0
# Author: Ermelino Piazzetta
# Creation Date: 2025-06-25
# Description: Updates the Excel summary with total project values.
# -------------------------------------------

import os
from openpyxl import Workbook, load_workbook

def update_project_summary(project_name: str, total: float) -> None:
    summary_file = "resumo_projetos.xlsx"

    if os.path.exists(summary_file):
        wb = load_workbook(summary_file)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Project Summary"
        ws.append(["Project", "Total Cost (R$)"])

    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[0].value == project_name:
            ws.delete_rows(row[0].row)