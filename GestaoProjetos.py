
# -*- coding: utf-8 -*-
# -------------------------------------------
# Project Management GUI - Version 2.1
# Author: Ermelino Piazzetta (modified)
# Using Tkinter with tabs and detailed views
# -------------------------------------------

import os
import platform
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.chart import BarChart, Reference
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()

def list_existing_projects():
    return [f[8:-5] for f in os.listdir() if f.startswith("project_") and f.endswith(".xlsx")]

def save_project_spreadsheet(project_name, items, totals_by_category):
    filename = f"project_{project_name}.xlsx"
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Project Items"
        ws.append(["Description", "Qty", "Unit", "Unit Price (R$)", "Total Price (R$)"])
    # Clear old data except header
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    start_row = 2
    for idx, item in enumerate(items, start=start_row):
        ws.cell(row=idx, column=1, value=item['desc'])
        ws.cell(row=idx, column=2, value=item['qty'])
        ws.cell(row=idx, column=3, value=item['unit'])
        ws.cell(row=idx, column=4, value=item['unit_price'])
        ws.cell(row=idx, column=5, value=item['total'])

    # Clear rows below items
    for row in ws.iter_rows(min_row=start_row+len(items), max_row=ws.max_row):
        for cell in row:
            cell.value = None

    totals_start = start_row + len(items) + 1
    ws.cell(row=totals_start, column=1, value="Totals by Category")
    ws.cell(row=totals_start+1, column=1, value="Category")
    ws.cell(row=totals_start+1, column=2, value="Total")

    for i, (cat, tot) in enumerate(totals_by_category.items(), start=totals_start+2):
        ws.cell(row=i, column=1, value=cat)
        ws.cell(row=i, column=2, value=tot)

    # Formatting
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="BDD7EE")
    border = Border(*(Side(style='thin'),)*4)
    money = NamedStyle(name="money", number_format='"R$"#,##0.00')
    if "money" not in wb.named_styles:
        wb.add_named_style(money)
    for r in ws.iter_rows(min_row=1, max_col=5, max_row=ws.max_row):
        for cell in r:
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if cell.row == 1:
                cell.font=bold; cell.fill=fill
            if cell.column in (4,5):
                cell.style="money"

    # Chart
    chart = BarChart()
    chart.title = "Totals per Category"
    chart.x_axis.title = "Category"
    chart.y_axis.title = "Total"

    data = Reference(ws, min_col=2, min_row=totals_start+1, max_row=ws.max_row)
    cats = Reference(ws, min_col=1, min_row=totals_start+2, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    # Remove old charts
    ws._charts.clear()
    ws.add_chart(chart, f"A{ws.max_row + 3}")

    wb.save(filename)
    try:
        if platform.system() == "Windows":
            os.startfile(filename)
        elif platform.system() == "Darwin":
            os.system(f"open {filename}")
        else:
            os.system(f"xdg-open {filename}")
    except:
        pass
    return sum(totals_by_category.values())

def save_project_info_sheet(project_name, info):
    filename = f"project_{project_name}.xlsx"
    wb = load_workbook(filename) if os.path.exists(filename) else Workbook()
    if "Information" in wb.sheetnames:
        wb.remove(wb["Information"])
    ws = wb.create_sheet(title="Information")
    ws.append(["Field", "Value"])
    for key in ["manager", "manager_email", "start_date", "end_date", "est_cost"]:
        ws.append([key, info.get(key, "")])
    ws.append([])
    ws.append(["Participant Name", "Email"])
    for p in info.get("participants", []):
        ws.append([p["name"], p["email"]])
    wb.save(filename)

def update_summary(project_name, total_cost):
    fn = "project_summary.xlsx"
    if os.path.exists(fn):
        wb = load_workbook(fn)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Summary"
        ws.append(["Project", "Total Cost"])
    # remove existing
    for row in ws.iter_rows(min_row=2):
        if row[0].value == project_name:
            ws.delete_rows(row[0].row)
    ws.append([project_name, total_cost])
    wb.save(fn)

def send_emails(project_name, info):
    from_addr = info.get('manager_email')
    pwd = os.getenv("EMAIL_PASSWORD")
    if not from_addr or not pwd:
        messagebox.showwarning("Email Warning", "Manager email or email password not configured.")
        return
    for p in info.get('participants', []):
        msg = EmailMessage()
        msg["Subject"] = f"New Project: {project_name}"
        msg["From"] = from_addr
        msg["To"] = p.get('email')
        msg.set_content(
            f"Hello {p.get('name')},\n\nYou were added to project \"{project_name}\".\n"
            f"Manager: {info.get('manager')}\nStart: {info.get('start_date')}\n"
            f"End: {info.get('end_date')}\nEstimated Cost: R$ {info.get('est_cost'):.2f}\n"
        )
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(from_addr, pwd)
                smtp.send_message(msg)
        except Exception as e:
            messagebox.showwarning("Email Error", f"Could not send to {p.get('email')}: {e}")

class ProjectManagerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Project Manager v2.0")
        self.geometry("700x500")

        self.projects = []
        self.project_info = None
        self.project_items = []

        self.create_menu()
        self.create_widgets()
        self.load_projects()

    def create_menu(self):
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Project", command=self.new_project_dialog)
        file_menu.add_command(label="Delete Project", command=self.delete_project)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        self.config(menu=menubar)

    def create_widgets(self):
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Left panel: project list
        left_frame = ttk.Frame(self.main_frame, width=200)
        left_frame.pack(side="left", fill="y")

        ttk.Label(left_frame, text="Projects").pack(pady=5)
        self.project_listbox = tk.Listbox(left_frame, exportselection=False)
        self.project_listbox.pack(fill="y", expand=True)
        self.project_listbox.bind("<<ListboxSelect>>", self.on_project_select)

        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="New Project", command=self.new_project_dialog).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Delete Project", command=self.delete_project).pack(side="left", padx=2)

        # Right panel: tabs for details
        right_frame = ttk.Frame(self.main_frame)
        right_frame.pack(side="left", fill="both", expand=True)

        self.tabs = ttk.Notebook(right_frame)
        self.tabs.pack(fill="both", expand=True)

        # Tab 1: Info
        self.tab_info = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_info, text="Project Info")

        self.info_tree = ttk.Treeview(self.tab_info, columns=("field", "value"), show="headings")
        self.info_tree.heading("field", text="Field")
        self.info_tree.heading("value", text="Value")
        self.info_tree.pack(fill="both", expand=True, padx=5, pady=5)

        # Tab 2: Participants
        self.tab_participants = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_participants, text="Participants")

        self.part_tree = ttk.Treeview(self.tab_participants, columns=("name", "email"), show="headings")
        self.part_tree.heading("name", text="Name")
        self.part_tree.heading("email", text="Email")
        self.part_tree.pack(fill="both", expand=True, padx=5, pady=5)

        # Tab 3: Items
        self.tab_items = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_items, text="Project Items")

        self.items_tree = ttk.Treeview(self.tab_items, columns=("desc", "qty", "unit", "unit_price", "total"), show="headings")
        self.items_tree.heading("desc", text="Description")
        self.items_tree.heading("qty", text="Quantity")
        self.items_tree.heading("unit", text="Unit")
        self.items_tree.heading("unit_price", text="Unit Price")
        self.items_tree.heading("total", text="Total Price")
        self.items_tree.pack(fill="both", expand=True, padx=5, pady=5)

        btn_items = ttk.Frame(self.tab_items)
        btn_items.pack(pady=5)
        ttk.Button(btn_items, text="Add Item", command=self.add_item_dialog).pack(side="left", padx=5)
        ttk.Button(btn_items, text="Remove Selected Item", command=self.remove_selected_item).pack(side="left", padx=5)
        ttk.Button(btn_items, text="Save Items", command=self.save_items).pack(side="left", padx=5)

    def load_projects(self):
        self.projects = list_existing_projects()
        self.project_listbox.delete(0, "end")
        if self.projects:
            for p in self.projects:
                self.project_listbox.insert("end", p)
        else:
            self.project_listbox.insert("end", "<No projects>")
        self.clear_detail_views()

    def clear_detail_views(self):
        for tree in (self.info_tree, self.part_tree, self.items_tree):
            for item in tree.get_children():
                tree.delete(item)
        self.project_info = None
        self.project_items = []

    def on_project_select(self, event):
        if not self.projects:
            return
        selection = self.project_listbox.curselection()
        if not selection:
            return
        project_name = self.project_listbox.get(selection[0])
        if project_name == "<No projects>":
            self.clear_detail_views()
            return
        self.load_project_details(project_name)

    def load_project_details(self, project_name):
        filename = f"project_{project_name}.xlsx"
        if not os.path.exists(filename):
            messagebox.showerror("Error", "Project file missing.")
            return

        wb = load_workbook(filename)

        # Load Information tab
        self.project_info = {}
        if "Information" in wb.sheetnames:
            ws_info = wb["Information"]
            self.info_tree.delete(*self.info_tree.get_children())
            for row in ws_info.iter_rows(min_row=2, values_only=True):
                if row[0] is None:
                    continue
                self.project_info[row[0]] = row[1]
                self.info_tree.insert("", "end", values=(row[0], row[1]))

        # Load participants
        self.part_tree.delete(*self.part_tree.get_children())
        participants = []
        if "Information" in wb.sheetnames:
            ws_info = wb["Information"]
            collecting = False
            for row in ws_info.iter_rows(values_only=True):
                if row and row[0] == "Participant Name":
                    collecting = True
                    continue
                if collecting and row and row[0]:
                    participants.append({"name": row[0], "email": row[1]})
            for p in participants:
                self.part_tree.insert("", "end", values=(p["name"], p["email"]))
        self.project_info["participants"] = participants

        # Load items
        self.project_items = []
        if "Project Items" in wb.sheetnames:
            ws = wb["Project Items"]
            self.items_tree.delete(*self.items_tree.get_children())
            for row in ws.iter_rows(min_row=2, max_col=5, values_only=True):
                if all(x is None for x in row):
                    continue
                if isinstance(row[0], str) and row[0].startswith("Totals by"):
                    break
                self.project_items.append({
                    "desc": row[0],
                    "qty": row[1],
                    "unit": row[2],
                    "unit_price": row[3],
                    "total": row[4],
                })
                self.items_tree.insert("", "end", values=row)

    def new_project_dialog(self):
        dlg = NewProjectDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            # dlg.result is a dict with info and items
            info = dlg.result['info']
            items = dlg.result['items']
            totals = defaultdict(float)
            for item in items:
                totals[item['desc']] += item['total']
            project_name = info.get('project_name')
            save_project_spreadsheet(project_name, items, totals)
            save_project_info_sheet(project_name, info)
            update_summary(project_name, sum(totals.values()))
            send_emails(project_name, info)
            self.load_projects()

    def delete_project(self):
        sel = self.project_listbox.curselection()
        if not sel or not self.projects:
            return
        project_name = self.project_listbox.get(sel[0])
        if messagebox.askyesno("Delete Project", f"Delete project '{project_name}'?"):
            fn = f"project_{project_name}.xlsx"
            try:
                os.remove(fn)
                update_summary(project_name, 0)
            except Exception as e:
                messagebox.showerror("Error", f"Could not delete: {e}")
            self.load_projects()

    def add_item_dialog(self):
        dlg = AddItemDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            # dlg.result = dict with item fields
            item = dlg.result
            # Check duplicate description
            if any(i['desc'].lower() == item['desc'].lower() for i in self.project_items):
                messagebox.showwarning("Duplicate", "Item with this description already exists.")
                return
            self.project_items.append(item)
            self.items_tree.insert("", "end", values=(item['desc'], item['qty'], item['unit'], item['unit_price'], item['total']))

    def remove_selected_item(self):
        sel = self.items_tree.selection()
        if not sel:
            return
        idx = self.items_tree.index(sel[0])
        self.items_tree.delete(sel[0])
        del self.project_items[idx]

    def save_items(self):
        if not self.project_info:
            messagebox.showwarning("Warning", "No project selected.")
            return
        project_name = self.project_info.get('project_name')
        if not project_name:
            messagebox.showwarning("Warning", "Invalid project info.")
            return
        totals = defaultdict(float)
        for item in self.project_items:
            totals[item['desc']] += item['total']
        save_project_spreadsheet(project_name, self.project_items, totals)
        update_summary(project_name, sum(totals.values()))
        messagebox.showinfo("Saved", "Project items saved.")

class NewProjectDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("New Project")
        self.result = None

        self.geometry("500x600")
        self.transient(parent)
        self.grab_set()

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        # Project name
        ttk.Label(frm, text="Project Name:").pack(anchor="w")
        self.entry_name = ttk.Entry(frm)
        self.entry_name.pack(fill="x")

        # Manager name/email
        ttk.Label(frm, text="Manager Name:").pack(anchor="w", pady=(10,0))
        self.entry_manager = ttk.Entry(frm)
        self.entry_manager.pack(fill="x")

        ttk.Label(frm, text="Manager Email:").pack(anchor="w", pady=(10,0))
        self.entry_manager_email = ttk.Entry(frm)
        self.entry_manager_email.pack(fill="x")

        # Dates
        ttk.Label(frm, text="Start Date (YYYY-MM-DD):").pack(anchor="w", pady=(10,0))
        self.entry_start = ttk.Entry(frm)
        self.entry_start.pack(fill="x")

        ttk.Label(frm, text="Estimated End Date:").pack(anchor="w", pady=(10,0))
        self.entry_end = ttk.Entry(frm)
        self.entry_end.pack(fill="x")

        # Estimated cost
        ttk.Label(frm, text="Estimated Cost (R$):").pack(anchor="w", pady=(10,0))
        self.entry_cost = ttk.Entry(frm)
        self.entry_cost.pack(fill="x")

        # Participants
        ttk.Label(frm, text="Participants (Add Name and Email):").pack(anchor="w", pady=(10,0))

        self.participants = []

        self.part_tree = ttk.Treeview(frm, columns=("name", "email"), show="headings", height=6)
        self.part_tree.heading("name", text="Name")
        self.part_tree.heading("email", text="Email")
        self.part_tree.pack(fill="both", pady=5, expand=True)

        btn_part = ttk.Frame(frm)
        btn_part.pack(fill="x")
        ttk.Button(btn_part, text="Add Participant", command=self.add_participant_dialog).pack(side="left", padx=5)
        ttk.Button(btn_part, text="Remove Selected", command=self.remove_selected_participant).pack(side="left", padx=5)

        # Project Items
        ttk.Label(frm, text="Project Items:").pack(anchor="w", pady=(10,0))

        self.items = []

        self.items_tree = ttk.Treeview(frm, columns=("desc", "qty", "unit", "unit_price", "total"), show="headings", height=6)
        self.items_tree.heading("desc", text="Description")
        self.items_tree.heading("qty", text="Quantity")
        self.items_tree.heading("unit", text="Unit")
        self.items_tree.heading("unit_price", text="Unit Price")
        self.items_tree.heading("total", text="Total Price")
        self.items_tree.pack(fill="both", pady=5, expand=True)

        btn_items = ttk.Frame(frm)
        btn_items.pack(fill="x")
        ttk.Button(btn_items, text="Add Item", command=self.add_item_dialog).pack(side="left", padx=5)
        ttk.Button(btn_items, text="Remove Selected", command=self.remove_selected_item).pack(side="left", padx=5)

        # Buttons
        btn_frame = ttk.Frame(frm)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Save Project", command=self.save_project).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="left", padx=10)

    def add_participant_dialog(self):
        dlg = AddParticipantDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            p = dlg.result
            # Check duplicate email
            if any(x['email'].lower() == p['email'].lower() for x in self.participants):
                messagebox.showwarning("Duplicate", "Participant with this email already exists.")
                return
            self.participants.append(p)
            self.part_tree.insert("", "end", values=(p["name"], p["email"]))

    def remove_selected_participant(self):
        sel = self.part_tree.selection()
        if not sel:
            return
        idx = self.part_tree.index(sel[0])
        self.part_tree.delete(sel[0])
        del self.participants[idx]

    def add_item_dialog(self):
        dlg = AddItemDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            item = dlg.result
            # Check duplicate description
            if any(i['desc'].lower() == item['desc'].lower() for i in self.items):
                messagebox.showwarning("Duplicate", "Item with this description already exists.")
                return
            self.items.append(item)
            self.items_tree.insert("", "end", values=(item['desc'], item['qty'], item['unit'], item['unit_price'], item['total']))

    def remove_selected_item(self):
        sel = self.items_tree.selection()
        if not sel:
            return
        idx = self.items_tree.index(sel[0])
        self.items_tree.delete(sel[0])
        del self.items[idx]

    def save_project(self):
        # Validate required fields
        name = self.entry_name.get().strip().replace(" ", "_")
        if not name:
            messagebox.showerror("Error", "Project Name is required.")
            return
        manager = self.entry_manager.get().strip()
        email = self.entry_manager_email.get().strip()
        start_date = self.entry_start.get().strip()
        end_date = self.entry_end.get().strip()
        try:
            est_cost = float(self.entry_cost.get().strip())
        except:
            messagebox.showerror("Error", "Estimated Cost must be a valid number.")
            return
        if not manager or not email or not start_date or not end_date:
            messagebox.showerror("Error", "All fields are required.")
            return
        if not self.items:
            messagebox.showerror("Error", "Add at least one project item.")
            return
        info = {
            "project_name": name,
            "manager": manager,
            "manager_email": email,
            "start_date": start_date,
            "end_date": end_date,
            "est_cost": est_cost,
            "participants": self.participants
        }
        self.result = {"info": info, "items": self.items}
        self.destroy()

class AddParticipantDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Add Participant")
        self.result = None
        self.geometry("300x150")
        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text="Name:").pack(anchor="w", padx=10, pady=(10,0))
        self.entry_name = ttk.Entry(self)
        self.entry_name.pack(fill="x", padx=10)

        ttk.Label(self, text="Email:").pack(anchor="w", padx=10, pady=(10,0))
        self.entry_email = ttk.Entry(self)
        self.entry_email.pack(fill="x", padx=10)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Add", command=self.on_add).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="left", padx=5)

    def on_add(self):
        name = self.entry_name.get().strip()
        email = self.entry_email.get().strip()
        if not name or not email:
            messagebox.showerror("Error", "Both name and email are required.")
            return
        self.result = {"name": name, "email": email}
        self.destroy()

class AddItemDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Add Item")
        self.result = None
        self.geometry("350x300")
        self.transient(parent)
        self.grab_set()

        frm = ttk.Frame(self)
        frm.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Label(frm, text="Description:").pack(anchor="w")
        self.entry_desc = ttk.Entry(frm)
        self.entry_desc.pack(fill="x", pady=2)

        ttk.Label(frm, text="Quantity:").pack(anchor="w", pady=(10,0))
        self.entry_qty = ttk.Entry(frm)
        self.entry_qty.pack(fill="x", pady=2)

        ttk.Label(frm, text="Unit:").pack(anchor="w", pady=(10,0))
        self.entry_unit = ttk.Entry(frm)
        self.entry_unit.pack(fill="x", pady=2)

        ttk.Label(frm, text="Unit Price (R$):").pack(anchor="w", pady=(10,0))
        self.entry_unit_price = ttk.Entry(frm)
        self.entry_unit_price.pack(fill="x", pady=2)

        btn_frame = ttk.Frame(frm)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Add", command=self.on_add).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="left", padx=5)

    def on_add(self):
        desc = self.entry_desc.get().strip()
        if not desc:
            messagebox.showerror("Error", "Description required.")
            return
        try:
            qty = float(self.entry_qty.get().strip())
            if qty <= 0:
                raise ValueError()
        except:
            messagebox.showerror("Error", "Quantity must be positive number.")
            return
        unit = self.entry_unit.get().strip()
        if not unit:
            messagebox.showerror("Error", "Unit is required.")
            return
        try:
            unit_price = float(self.entry_unit_price.get().strip())
            if unit_price < 0:
                raise ValueError()
        except:
            messagebox.showerror("Error", "Unit price must be non-negative number.")
            return
        total = qty * unit_price
        self.result = {
            "desc": desc,
            "qty": qty,
            "unit": unit,
            "unit_price": unit_price,
            "total": total
        }
        self.destroy()

if __name__ == "__main__":
    app = ProjectManagerApp()
    app.mainloop()
