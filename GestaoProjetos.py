# -*- coding: utf-8 -*-
# -------------------------------------------
# Project Management - Version 1.0
# Author: Ermelino Piazzetta
# Creation Date: 2025-06-25
# Description: All-in-one project management system script.
# -------------------------------------------

import os
import platform
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.chart import BarChart, Reference

# === Input Utilities ===

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
    totals_by_category = defaultdict(float)

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

        totals_by_category[description] += total_price
        print("Item successfully added!\n")

    return items, totals_by_category

# === File Utilities ===

def open_file(filename: str):
    system = platform.system()
    if system == "Windows":
        os.startfile(filename)
    elif system == "Darwin":
        os.system(f"open {filename}")
    else:
        os.system(f"xdg-open {filename}")

def list_existing_projects():
    return [f[8:-5] for f in os.listdir() if f.startswith("projeto_") and f.endswith(".xlsx")]

# === Spreadsheet Utilities ===

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
    open_file(filename)

    return sum(totals_by_category.values())

def save_project_info_sheet(project_name: str, info: dict) -> None:
    filename = f"projeto_{project_name}.xlsx"
    if os.path.exists(filename):
        wb = load_workbook(filename)
    else:
        wb = Workbook()

    if "Project Info" in wb.sheetnames:
        wb.remove(wb["Project Info"])

    ws = wb.create_sheet(title="Project Info")
    ws.append(["Field", "Value"])
    for key in ["Manager", "Opening Date", "Estimated Completion", "Estimated Cost"]:
        ws.append([key, info[key]])

    ws.append([])
    ws.append(["Participant Name", "Email", "Phone Number"])
    for participant in info.get("Participants", []):
        ws.append([participant["Name"], participant["Email"], participant["Phone"]])

    wb.save(filename)
    print(f"Project info sheet updated in '{filename}'.")

def update_project_summary(project_name: str, project_total: float):
    summary_filename = "resumo_projetos.xlsx"

    if os.path.exists(summary_filename):
        wb = load_workbook(summary_filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Project Summary"
        ws.append(["Project", "Total Cost (R$)"])

    # Remove existing row for the project if it exists
    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[0].value == project_name:
            ws.delete_rows(row[0].row)

    ws.append([project_name, project_total])
    wb.save(summary_filename)
    print(f"Summary updated in '{summary_filename}'.")

def main():
    print("=== Project Registration System ===")

    while True:
        choice = input("\nDo you want to start a (N)ew project or open an (E)xisting one? (n/e): ").strip().lower()
        if choice == 'n':
            project_name = input("Enter the new project name: ").strip().replace(" ", "_")
            # Gather project info
            manager = input("Project Manager Name: ").strip()
            opening_date = input("Opening Date (YYYY-MM-DD): ").strip()
            estimated_completion = input("Estimated Completion Date (YYYY-MM-DD): ").strip()
            estimated_cost = prompt_for_value("Estimated Cost (R$): ", float)

            participants = []
            print("\nEnter project participants (leave name empty to finish):")
            while True:
                name = input("Participant Name: ").strip()
                if not name:
                    break
                email = input("Participant Email: ").strip()
                phone = input("Participant Phone: ").strip()
                participants.append({"Name": name, "Email": email, "Phone": phone})

            project_info = {
                "Manager": manager,
                "Opening Date": opening_date,
                "Estimated Completion": estimated_completion,
                "Estimated Cost": estimated_cost,
                "Participants": participants
            }
            save_project_info_sheet(project_name, project_info)
            break

        elif choice == 'e':
            projects = list_existing_projects()
            if not projects:
                print("No projects found.")
                return
            print("Existing projects:")
            for i, p in enumerate(projects, 1):
                print(f"{i}. {p}")
            selection = prompt_for_value("Select the project number: ", int)
            if 1 <= selection <= len(projects):
                project_name = projects[selection - 1]
                break
            else:
                print("Invalid choice.")
        else:
            print("Invalid option. Please type 'n' or 'e'.")

    while True:
        items, totals = register_items()
        if items:
            total_cost = save_project_spreadsheet(project_name, items, totals)
            update_project_summary(project_name, total_cost)
        else:
            print("No items were registered.")

        cont = input("\nDo you want to register more items in this project? (y/n): ").strip().lower()
        if cont != 'y':
            print("Ending registration for this project.")
            break

if __name__ == "__main__":
    main()