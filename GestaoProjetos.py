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

def criar_novo_arquivo(nome_base: str = "itens_projeto") -> str:
    contador = 1
    while True:
        nome_arquivo = f"{nome_base}_{contador}.xlsx"
        if not os.path.exists(nome_arquivo):
            return nome_arquivo
        contador += 1

def aplicar_formatacao(ws):
    # Define estilos
    bold_font = Font(bold=True)
    fill_header = PatternFill("solid", fgColor="BDD7EE")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    money_style = NamedStyle(name="money_style", number_format='"R$"#,##0.00')
    if "money_style" not in ws.parent.named_styles:
        ws.parent.add_named_style(money_style)

    # Aplicar estilos nos dados principais
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=5):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.row == 1:  # Cabeçalho
                cell.font = bold_font
                cell.fill = fill_header
            elif cell.column in (4, 5):  # Colunas de valor
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
    print("\n=== Cadastro de Itens do Projeto ===")
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

def salvar_em_planilha(itens, totais_por_categoria, nome_arquivo="itens_projeto.xlsx", nova_planilha=False):
    if nova_planilha or not os.path.exists(nome_arquivo):
        wb = Workbook()
        ws = wb.active
        ws.title = "Itens do Projeto"
        ws.append(["Descrição", "Quantidade", "Unidade", "Valor Unitário (R$)", "Valor Total (R$)"])
    else:
        wb = load_workbook(nome_arquivo)
        ws = wb.active

    escrever_itens_na_planilha(ws, itens, totais_por_categoria)
    wb.save(nome_arquivo)
    print(f"\nArquivo '{nome_arquivo}' salvo com sucesso.")
    abrir_planilha(nome_arquivo)

def main():
    nome_padrao = "itens_projeto.xlsx"

    while True:
        itens, totais = cadastrar_itens()

        if not itens:
            print("Nenhum item foi cadastrado.")
        else:
            if os.path.exists(nome_padrao):
                print("\n⚙️ A planilha padrão já existe.")
                escolha = input("Deseja (A)dicionar à planilha existente, (N)ova planilha, ou (S)air? ").strip().lower()

                if escolha == 'a':
                    salvar_em_planilha(itens, totais, nome_padrao, nova_planilha=False)
                elif escolha == 'n':
                    novo_arquivo = criar_novo_arquivo("itens_projeto")
                    salvar_em_planilha(itens, totais, nome_arquivo=novo_arquivo, nova_planilha=True)
                elif escolha == 's':
                    print("Encerrando o programa.")
                    break
                else:
                    print("Opção inválida. Encerrando.")
                    break
            else:
                salvar_em_planilha(itens, totais, nome_padrao, nova_planilha=True)

        continuar = input("\nDeseja continuar cadastrando itens? (s/n): ").strip().lower()
        if continuar != 's':
            print("Programa finalizado.")
            break

if __name__ == "__main__":
    main()