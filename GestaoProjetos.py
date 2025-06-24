import openpyxl
from openpyxl import Workbook
import os
import platform
from collections import defaultdict

def solicitar_dado(mensagem: str, tipo: type):
    """
    Solicita um dado do tipo especificado ao usuário.
    """
    while True:
        entrada = input(mensagem)
        try:
            return tipo(entrada)
        except ValueError:
            print(f"Entrada inválida. Por favor, insira um valor do tipo {tipo.__name__}.")

def abrir_planilha(nome_arquivo: str):
    """
    Abre a planilha Excel com o programa padrão do sistema.
    """
    if platform.system() == "Windows":
        os.startfile(nome_arquivo)
    elif platform.system() == "Darwin":  # macOS
        os.system(f"open {nome_arquivo}")
    else:  # Linux
        os.system(f"xdg-open {nome_arquivo}")

def main():
    print("=== Cadastro de Itens do Projeto ===")
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

    if not itens:
        print("Nenhum item foi cadastrado. Programa encerrado.")
        return

    # Criar planilha Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Itens do Projeto"

    # Cabeçalho
    ws.append(["Descrição", "Quantidade", "Unidade", "Valor Unitário (R$)", "Valor Total (R$)"])

    # Inserir dados
    for item in itens:
        ws.append([
            item["Descrição"],
            item["Quantidade"],
            item["Unidade"],
            item["Valor Unitário"],
            item["Valor Total"]
        ])

    # Linha em branco e totais por categoria
    ws.append([])
    ws.append(["Totais por Categoria"])
    ws.append(["Categoria", "Total (R$)"])
    for categoria, total in totais_por_categoria.items():
        ws.append([categoria, total])

    # Salvar e abrir arquivo
    nome_arquivo = "itens_projeto.xlsx"
    wb.save(nome_arquivo)
    print(f"\nArquivo '{nome_arquivo}' salvo com sucesso.")

    abrir_planilha(nome_arquivo)

if __name__ == "__main__":
    main()