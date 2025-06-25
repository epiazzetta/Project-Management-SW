
# -*- main.py -*-
# -------------------------------------------
# Project Management    - Version 1.0
# Author: Ermelino Piazzetta
# Creation Date: 2025-06-25
# Description: Main entry point of the system.
# -------------------------------------------

from input_utils import prompt_for_value, register_items
from file_utils import open_file, list_existing_projects
from spreadsheet import save_project_spreadsheet
from summary import update_project_summary

def main() -> None:
    print("=== Project Management    ===")

    while True:
        choice = input("\nStart a (N)ew project or open an (E)xisting one? (n/e): ").strip().lower()
        if choice == 'n':
            project_name = input("Enter the new project name: ").strip().replace(" ", "_")
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

from typing import Any, Tuple, List, Dict
from collections import defaultdict

def prompt_for_value(message: str, value_type: type) -> Any:
    while True:
        user_input = input(message)
        try:
            return value_type(user_input)
        except ValueError:
            print(f"Invalid input. Expected a {value_type.__name__}.")

def register_items() -> Tuple[List[Dict], Dict[str, float]]:
    print("\n=== Item Registration ===")
    print("Type 'end' as description to finish.\n")

    items = []
    totals = defaultdict(float)

    while True:
        desc = input("Item description (or 'end'): ").strip()
        if desc.lower() == 'end':
            break
        quantity = prompt_for_value("Quantity: ", float)
        unit = input("Unit (e.g., hour, material): ").strip()
        unit_price = prompt_for_value("Unit price (R$): ", float)
        total_price = quantity * unit_price

        items.append({
            "Description": desc,
            "Quantity": quantity,
            "Unit": unit,
            "Unit Price": unit_price,
            "Total": total_price
        })
        totals[desc] += total_price
        print("Item added!\n")

    return items, totals

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
from typing import List, Dict
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.chart import BarChart, Reference

from file_utils import open_file

def apply_formatting(ws: Worksheet) -> None:
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="BDD7EE")
    border = Border(*(Side(style="thin") for _ in range(4)))

    money_style = NamedStyle(name="money", number_format='"R$"#,##0.00')
    if "money" not in ws.parent.named_styles:
        ws.parent.add_named_style(money_style)

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=5):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.row == 1:
                cell.font = bold
                cell.fill = fill
            elif cell.column in (4, 5):
                cell.style = money_style

def add_chart(ws: Worksheet, start_row: int) -> None:
    chart = BarChart()
    chart.title = "Total per Category"
    chart.x_axis.title = "Category"
    chart.y_axis.title = "Total Value (R$)"

    data = Reference(ws, min_col=2, min_row=start_row + 1, max_row=ws.max_row)
    categories = Reference(ws, min_col=1, min_row=start_row + 2, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    ws.add_chart(chart, f"A{ws.max_row + 3}")

def write_items(ws: Worksheet, items: List[Dict], totals: Dict[str, float]) -> None:
    for item in items:
        ws.append([
            item["Description"],
            item["Quantity"],
            item["Unit"],
            item["Unit Price"],
            item["Total"]
        ])
    ws.append([])
    start_row = ws.max_row + 1
    ws.append(["Total per Category"])
    ws.append(["Category", "Total (R$)"])
    for category, total in totals.items():
        ws.append([category, total])

    apply_formatting(ws)
    add_chart(ws, start_row)

def save_project_spreadsheet(project_name: str, items: List[Dict], totals: Dict[str, float]) -> float:
    file_name = f"projeto_{project_name}.xlsx"

    if os.path.exists(file_name):
        wb = load_workbook(file_name)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Project Items"
        ws.append(["Description", "Quantity", "Unit", "Unit Price (R$)", "Total (R$)"])

    write_items(ws, items, totals)
    wb.save(file_name)
    print(f"\nFile '{file_name}' saved successfully.")
    open_file(file_name)

    return sum(totals.values())

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
            ws.delete_rows(row[0].row