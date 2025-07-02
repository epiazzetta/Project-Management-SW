

# -*- coding: utf-8 -*-
# -------------------------------------------
# Project Management - Versão 1.5
# Autor: Ermelino Piazzetta (modificado por segurança)
# -------------------------------------------

import os
import platform
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.chart import BarChart, Reference
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

# Carrega variáveis de ambiente
load_dotenv()


def get_validated_input(prompt: str, convert_func=str, capitalize=False):
    while True:
        value = input(prompt).strip()
        if value.lower() == "end":
            return "end"
        if convert_func == str and capitalize:
            value = value.title()
        try:
            value = convert_func(value)
        except ValueError:
            print(f"Entrada inválida. Esperado tipo {convert_func.__name__}.")
            continue
        print(f"Você digitou: {value}")
        confirm = input("Confirma esta informação? (s/n): ").strip().lower()
        if confirm == 's':
            return value


def register_items():
    print("\n=== Item Registration ===")
    print("Digite 'end' para encerrar.\n")
    items = []
    totals_by_category = defaultdict(float)

    while True:
        description = get_validated_input("Descrição: ", str, capitalize=True)
        if description.lower() == 'end':
            break
        quantity = get_validated_input("Quantidade: ", float)
        unit = get_validated_input("Unidade (ex: hora, metro): ")
        unit_price = get_validated_input("Preço unitário (R$): ", float)
        total_price = quantity * unit_price

        items.append({
            "Description": description,
            "Quantity": quantity,
            "Unit": unit,
            "Unit Price": unit_price,
            "Total Price": total_price
        })

        totals_by_category[description] += total_price
        print("Item adicionado!\n")

    return items, totals_by_category


def open_file(filename: str):
    if platform.system() == "Windows":
        os.startfile(filename)
    elif platform.system() == "Darwin":
        os.system(f"open {filename}")
    else:
        os.system(f"xdg-open {filename}")


def list_existing_projects():
    return [f[8:-5] for f in os.listdir() if f.startswith("projeto_") and f.endswith(".xlsx")]


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
    chart.title = "Totais por Categoria"
    chart.x_axis.title = "Categoria"
    chart.y_axis.title = "Total (R$)"

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
    ws.append(["Totais por Categoria"])
    ws.append(["Categoria", "Total (R$)"])
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
        ws.title = "Itens do Projeto"
        ws.append(["Descrição", "Quantidade", "Unidade", "Preço Unitário (R$)", "Preço Total (R$)"])

    write_items_to_sheet(ws, items, totals_by_category)
    wb.save(filename)
    print(f"\nArquivo '{filename}' salvo com sucesso.")
    open_file(filename)

    return sum(totals_by_category.values())


def save_project_info_sheet(project_name: str, info: dict):
    filename = f"projeto_{project_name}.xlsx"
    wb = load_workbook(filename) if os.path.exists(filename) else Workbook()

    if "Informações" in wb.sheetnames:
        wb.remove(wb["Informações"])
    ws = wb.create_sheet(title="Informações")
    ws.append(["Campo", "Valor"])
    for key in ["Manager", "ManagerEmail", "Opening Date", "Estimated Completion", "Estimated Cost"]:
        ws.append([key, info[key]])

    ws.append([])
    ws.append(["Nome Participante", "Email", "Telefone"])
    for p in info.get("Participants", []):
        ws.append([p["Name"], p["Email"], p["Phone"]])

    wb.save(filename)
    print(f"Informações do projeto salvas em '{filename}'.")


def update_project_summary(project_name: str, project_total: float):
    summary_filename = "resumo_projetos.xlsx"
    if os.path.exists(summary_filename):
        wb = load_workbook(summary_filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Resumo"
        ws.append(["Projeto", "Custo Total (R$)"])

    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[0].value == project_name:
            ws.delete_rows(row[0].row)
    ws.append([project_name, project_total])
    wb.save(summary_filename)
    print(f"Resumo atualizado em '{summary_filename}'.")


def send_project_email(project_name, info):
    from_email = info["ManagerEmail"]
    app_password = os.getenv("EMAIL_PASSWORD")

    for participant in info.get("Participants", []):
        msg = EmailMessage()
        msg["Subject"] = f"Abertura do projeto: {project_name}"
        msg["From"] = from_email
        msg["To"] = participant["Email"]
        msg.set_content(
            f"""
Olá {participant['Name']},

Você foi registrado como participante do projeto "{project_name}".

Gerente do Projeto: {info['Manager']}
Data de Abertura: {info['Opening Date']}
Conclusão Estimada: {info['Estimated Completion']}
Custo Estimado: R$ {info['Estimated Cost']:.2f}

Obrigado.
""".strip()
        )

        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(from_email, app_password)
                smtp.send_message(msg)
            print(f"E-mail enviado para {participant['Email']}")
        except Exception as e:
            print(f"Erro ao enviar e-mail para {participant['Email']}: {e}")


def main():
    print("=== Sistema de Registro de Projetos ===")

    while True:
        choice = input("\nNovo projeto (n) ou existente (e)? ").strip().lower()
        if choice == 'n':
            project_name = get_validated_input("Nome do projeto: ", str, capitalize=True).replace(" ", "_")
            manager = get_validated_input("Nome do gerente: ", str, capitalize=True)
            manager_email = get_validated_input("Email do gerente (remetente): ", str)
            opening_date = get_validated_input("Data de abertura (YYYY-MM-DD): ", str)
            estimated_completion = get_validated_input("Data estimada de conclusão (YYYY-MM-DD): ", str)
            estimated_cost = get_validated_input("Custo estimado (R$): ", float)

            participants = []
            print("\nCadastro de participantes (deixe nome em branco para terminar):")
            while True:
                name = input("Nome: ").strip()
                if not name:
                    break
                name = name.title()
                email = input("Email: ").strip()
                phone = input("Telefone: ").strip()
                participants.append({"Name": name, "Email": email, "Phone": phone})

            project_info = {
                "Manager": manager,
                "ManagerEmail": manager_email,
                "Opening Date": opening_date,
                "Estimated Completion": estimated_completion,
                "Estimated Cost": estimated_cost,
                "Participants": participants
            }

            save_project_info_sheet(project_name, project_info)
            send_project_email(project_name, project_info)
            break

        elif choice == 'e':
            projects = list_existing_projects()
            if not projects:
                print("Nenhum projeto encontrado.")
                return
            for i, name in enumerate(projects, 1):
                print(f"{i}. {name}")
            selection = get_validated_input("Escolha o número do projeto: ", int)
            if 1 <= selection <= len(projects):
                project_name = projects[selection - 1]
                break
            else:
                print("Opção inválida.")
        else:
            print("Opção inválida. Use 'n' ou 'e'.")

    while True:
        print("\nDeseja adicionar novos itens ao projeto?")
        cont = input("Digite 's' para sim ou qualquer outra tecla para encerrar: ").strip().lower()
        if cont != 's':
            print("Encerrando o projeto.")
            break

        items, totals = register_items()
        if items:
            total_cost = save_project_spreadsheet(project_name, items, totals)
            update_project_summary(project_name, total_cost)
        else:
            print("Nenhum item registrado.")



if __name__ == "__main__":
    main()