//version 1.3
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.chart import BarChart, Reference
import os
import platform
from collections import defaultdict

def solicitar_dado(mensagem: str, tipo: type):
    while True:
        entrada = input(mensagem)
        try:
            return tipo(entrada)
        except ValueError:
            print(f"Entrada inválida. Por favor, insira um valor do tipo {tipo.__name__}.")

def abrir_planilha(nome_arquivo: str):
    if platform.system() == "Windows":
        os.startfile(nome_arquivo)
    elif platform.system() == "Darwin":
        os.system(f"open {nome_arquivo}")
    else:
        os.system(f"xdg-open {nome_arquivo}")

def aplicar_formatacao(ws):
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

def adicionar_grafico(ws, linha_inicio: int):
    chart = BarChart()
    chart.title = "Totais por Categoria"
    chart.x_axis.title = "Categoria"
    chart.y_axis.title = "Valor Total (R$)"

    dados = Reference(ws, min_col=2, min_row=linha_inicio + 1, max_row=ws.max_row)
    categorias = Reference(ws, min_col=1, min_row=linha_inicio + 2, max_row=ws.max_row)

    chart.add_data(dados, titles_from_data=True)
    chart.set_categories(categorias)

    ws.add_chart(chart, f"A{ws.max_row + 3}")

def escrever_itens_na_planilha(ws, itens, totais_por_categoria):
    for item in itens:
        ws.append([
            item["Descrição"],
            item["Quantidade"],
            item["Unidade"],
            item["Valor Unitário"],
            item["Valor Total"]
        ])
    ws.append([])
    linha_inicio_totais = ws.max_row + 1
    ws.append(["Totais por Categoria"])
    ws.append(["Categoria", "Total (R$)"])
    for categoria, total in totais_por_categoria.items():
        ws.append([categoria, total])

    aplicar_formatacao(ws)
    adicionar_grafico(ws, linha_inicio_totais)

def cadastrar_itens():
    print("\n=== Cadastro de Itens ===")
    print("Digite 'fim' na descrição para encerrar a entrada de dados.\n")

    itens = []
    totais_por_categoria = defaultdict(float)

    while True:
        descricao = input("Descrição do item (ou 'fim' para encerrar): ").strip()
        if descricao.lower() == 'fim':
            break

        quantidade = solicitar_dado("Quantidade: ", float)
        unidade = input("Unidade (ex: homem hora, material): ").strip()
        valor_unitario = solicitar_dado("Valor unitário (R$): ", float)
        valor_total = quantidade * valor_unitario

        itens.append({
            "Descrição": descricao,
            "Quantidade": quantidade,
            "Unidade": unidade,
            "Valor Unitário": valor_unitario,
            "Valor Total": valor_total
        })

        totais_por_categoria[descricao] += valor_total
        print("Item adicionado com sucesso!\n")

    return itens, totais_por_categoria

def salvar_em_planilha_projeto(nome_projeto: str, itens, totais_por_categoria):
    nome_arquivo = f"projeto_{nome_projeto}.xlsx"
    if os.path.exists(nome_arquivo):
        wb = load_workbook(nome_arquivo)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Itens do Projeto"
        ws.append(["Descrição", "Quantidade", "Unidade", "Valor Unitário (R$)", "Valor Total (R$)"])

    escrever_itens_na_planilha(ws, itens, totais_por_categoria)
    wb.save(nome_arquivo)
    print(f"\nArquivo '{nome_arquivo}' salvo com sucesso.")
    abrir_planilha(nome_arquivo)

    return sum(totais_por_categoria.values())

def atualizar_resumo_projetos(nome_projeto: str, total_projeto: float):
    nome_resumo = "resumo_projetos.xlsx"

    if os.path.exists(nome_resumo):
        wb = load_workbook(nome_resumo)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Resumo de Projetos"
        ws.append(["Projeto", "Gasto Total (R$)"])

    # Remover linha antiga se já existe
    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[0].value == nome_projeto:
            ws.delete_rows(row[0].row)

    ws.append([nome_projeto, total_projeto])
    wb.save(nome_resumo)
    print(f"Resumo atualizado em '{nome_resumo}'.")

def listar_projetos_existentes():
    return [f[8:-5] for f in os.listdir() if f.startswith("projeto_") and f.endswith(".xlsx")]

def main():
    print("=== Sistema de Cadastro de Projetos ===")
    while True:
        tipo = input("\nVocê deseja iniciar um (N)ovo projeto ou abrir um (E)xistente? (n/e): ").strip().lower()
        if tipo == 'n':
            nome_projeto = input("Digite o nome do novo projeto: ").strip().replace(" ", "_")
            break
        elif tipo == 'e':
            projetos = listar_projetos_existentes()
            if not projetos:
                print("Nenhum projeto encontrado.")
                return
            print("Projetos existentes:")
            for i, p in enumerate(projetos, start=1):
                print(f"{i}. {p}")
            escolha = solicitar_dado("Escolha o número do projeto: ", int)
            if 1 <= escolha <= len(projetos):
                nome_projeto = projetos[escolha - 1]
                break
            else:
                print("Escolha inválida.")
        else:
            print("Opção inválida. Digite 'n' ou 'e'.")

    while True:
        itens, totais = cadastrar_itens()
        if itens:
            total_do_projeto = salvar_em_planilha_projeto(nome_projeto, itens, totais)
            atualizar_resumo_projetos(nome_projeto, total_do_projeto)
        else:
            print("Nenhum item foi cadastrado.")

        continuar = input("\nDeseja cadastrar mais itens neste projeto? (s/n): ").strip().lower()
        if continuar != 's':
            print("Encerrando o cadastro para este projeto.")
            break

if __name__ == "__main__":
    main()