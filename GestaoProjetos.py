

# -*- coding: utf-8 -*-
# -------------------------------------------
# Project Management GUI - Version 1.0
# Author: Ermelino Piazzetta (modified)
# Using Tkinter for GUI
# -------------------------------------------

import os
import platform
import tkinter as tk
from tkinter import simpledialog, messagebox, ttk, filedialog
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
    # write
    start_row = ws.max_row
    for item in items:
        ws.append([item['desc'], item['qty'], item['unit'], item['unit_price'], item['total']])
    ws.append([])
    totals_start = ws.max_row
    ws.append(["Totals by Category"])
    ws.append(["Category", "Total"])
    for cat, tot in totals_by_category.items():
        ws.append([cat, tot])
    # formatting
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
    # chart
    chart = BarChart()
    chart.title = "Totals per Category"
    chart.x_axis.title = "Category"; chart.y_axis.title="Total"
    data = Reference(ws, min_col=2, min_row=totals_start+1, max_row=ws.max_row)
    cats = Reference(ws, min_col=1, min_row=totals_start+2, max_row=ws.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, f"A{ws.max_row+3}")
    wb.save(filename)
    try:
        if platform.system()=="Windows": os.startfile(filename)
        elif platform.system()=="Darwin": os.system(f"open {filename}")
        else: os.system(f"xdg-open {filename}")
    except: pass
    return sum(totals_by_category.values())

def update_summary(project_name, total_cost):
    fn="project_summary.xlsx"
    if os.path.exists(fn):
        wb=load_workbook(fn); ws=wb.active
    else:
        wb=Workbook(); ws=wb.active
        ws.title="Summary"
        ws.append(["Project","Total Cost"])
    # remove existing
    for row in ws.iter_rows(min_row=2):
        if row[0].value==project_name:
            ws.delete_rows(row[0].row)
    ws.append([project_name, total_cost])
    wb.save(fn)

def send_emails(project_name, info):
    from_addr = info['manager_email']
    pwd = os.getenv("EMAIL_PASSWORD")
    for p in info['participants']:
        msg=EmailMessage()
        msg["Subject"] = f"New Project: {project_name}"
        msg["From"] = from_addr
        msg["To"] = p['email']
        msg.set_content(f"Hello {p['name']},\n\nYou were added to project \"{project_name}\".\nManager: {info['manager']}\nStart: {info['start_date']}\nEnd: {info['end_date']}\nEstimated Cost: R$ {info['est_cost']:.2f}")
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com",465) as smtp:
                smtp.login(from_addr,pwd)
                smtp.send_message(msg)
        except:
            messagebox.showwarning("Email Error", f"Could not send to {p['email']}")

class ProjectManagerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Project Manager")
        self.geometry("400x300")
        self.projects = []
        self.create_widgets()
        self.refresh_projects()

    def create_widgets(self):
        self.lst = tk.Listbox(self)
        self.lst.pack(fill="both", expand=True, padx=10, pady=5)

        fr = tk.Frame(self)
        fr.pack(pady=5)
        for txt,cmd in [("New Project",self.new_project),("Edit Project",self.edit_project),("Delete Project",self.delete_project),("Exit",self.quit)]:
            tk.Button(fr,text=txt,command=cmd).pack(side="left", padx=5)

    def refresh_projects(self):
        self.projects = list_existing_projects()
        self.lst.delete(0,"end")
        if self.projects:
            for p in self.projects:
                self.lst.insert("end", p)
        else:
            self.lst.insert("end","<No projects>")

    def new_project(self):
        name = simpledialog.askstring("New Project","Project Name:")
        if not name: return
        name = name.strip().replace(" ","_")
        if name in self.projects:
            messagebox.showerror("Error","Project exists")
            return
        manager = simpledialog.askstring("Manager Info","Manager Name:")
        email = simpledialog.askstring("Manager Email","Manager Email:")
        sd = simpledialog.askstring("Dates","Start Date (YYYY-MM-DD):")
        ed = simpledialog.askstring("Dates","Estimated End Date:")
        cost = simpledialog.askfloat("Cost","Estimated Cost (R$):")
        participants = []
        while True:
            pn = simpledialog.askstring("Participant","Name (blank to finish):")
            if not pn: break
            pe = simpledialog.askstring("Participant","Email:")
            participants.append({"name":pn,"email":pe})
        info = {"manager":manager,"manager_email":email,"start_date":sd,"end_date":ed,"est_cost":cost,"participants":participants}
        # Items
        items=[]; totals=defaultdict(float)
        while True:
            desc = simpledialog.askstring("Item","Description (blank to finish):")
            if not desc: break
            qty = simpledialog.askfloat("Item","Quantity:")
            unit = simpledialog.askstring("Item","Unit:")
            up = simpledialog.askfloat("Item","Unit Price (R$):")
            total = qty * up
            items.append({"desc":desc,"qty":qty,"unit":unit,"unit_price":up,"total":total})
            totals[desc]+=total
        tc = save_project_spreadsheet(name, items, totals)
        update_summary(name, tc)
        send_emails(name, info)
        self.refresh_projects()

    def edit_project(self):
        sel=self.lst.curselection()
        if not sel or not self.projects: return
        name=self.projects[sel[0]]
        items=[]; totals=defaultdict(float)
        while True:
            desc = simpledialog.askstring("Add Item","Description (blank to finish):")
            if not desc: break
            qty = simpledialog.askfloat("Qty","Quantity:")
            unit = simpledialog.askstring("Unit","Unit:")
            up = simpledialog.askfloat("Unit Price","Unit Price (R$):")
            total = qty * up
            items.append({"desc":desc,"qty":qty,"unit":unit,"unit_price":up,"total":total})
            totals[desc]+=total
        if items:
            tc = save_project_spreadsheet(name, items, totals)
            update_summary(name, tc)
            messagebox.showinfo("Updated", f"Added items to {name}")

    def delete_project(self):
        sel=self.lst.curselection()
        if not sel or not self.projects: return
        name=self.projects[sel[0]]
        if messagebox.askyesno("Delete","Delete project "+name+"?"):
            fn=f"project_{name}.xlsx"
            os.remove(fn)
            update_summary(name,0)
            self.refresh_projects()

if __name__ == "__main__":
    app = ProjectManagerApp()
    app.mainloop()
