"""
Script para automatizar o fluxo:
1. Abrir a planilha de origem (com Power Query)
2. Atualizar as consultas (equivalente ao Alt+F5)
3. Exportar os dados atualizados para dados.csv
4. Subir o CSV atualizado para o GitHub (add, commit, push)

Requisitos:
    pip install pywin32

Rode este script no Windows (não funciona em Mac/Linux, pois depende do Excel via COM).
"""

import time
import subprocess
import win32com.client as win32

# ===================== CONFIGURAÇÕES =====================
# Caminho completo da planilha de origem (ajuste a extensão se for .xlsm em vez de .xlsx)
CAMINHO_PLANILHA_ORIGEM = r"C:\Users\jaildo.junior\Desktop\DASHBOARD_RFK\DADOS_INSUMO\TESTE_ABC_DASHBOARD_1.xlsx"

# Nome da aba/planilha que contém a tabela final que deve virar o dados.csv
NOME_ABA_TABELA = "BASE DE DADOS"

# Caminho de saída do CSV
CAMINHO_SAIDA_CSV = r"C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK\dados.csv"

# Pasta raiz do repositório Git (a pasta que contém a pasta .git)
# >>> CONFIRME se é essa mesma pasta <<<
CAMINHO_REPO_GIT = r"C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK"

# Nome do arquivo dentro do repositório para o git add (relativo ao CAMINHO_REPO_GIT)
ARQUIVO_NO_REPO = "dados.csv"

MENSAGEM_COMMIT = "Atualização automática dos dados"
# ===========================================================


def atualizar_e_exportar():
    # DispatchEx força a criação de uma instância NOVA e isolada do Excel,
    # em vez de reaproveitar uma instância já aberta (que poderia estar visível).
    # Isso não fecha nem interfere em outras planilhas que você já tenha aberto manualmente.
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        print("Abrindo planilha de origem...")
        wb = excel.Workbooks.Open(CAMINHO_PLANILHA_ORIGEM)

        print("Atualizando Power Query (RefreshAll)...")
        wb.RefreshAll()

        # Consultas do Power Query rodam em segundo plano (assíncronas).
        # Isso força o Excel a esperar até todas terminarem.
        excel.CalculateUntilAsyncQueriesDone()

        # Pequena margem de segurança extra
        time.sleep(3)

        print(f"Exportando aba '{NOME_ABA_TABELA}' como CSV...")
        aba = wb.Worksheets(NOME_ABA_TABELA)
        aba.Copy()  # cria um novo workbook temporário só com essa aba
        novo_wb = excel.ActiveWorkbook
        novo_wb.SaveAs(CAMINHO_SAIDA_CSV, FileFormat=62)  # 62 = CSV UTF-8 (preserva acentos como em "MÊS")
        novo_wb.Close(SaveChanges=False)

        wb.Close(SaveChanges=False)
        print(f"CSV salvo com sucesso em: {CAMINHO_SAIDA_CSV}")
    finally:
        excel.Quit()


def subir_para_github():
    print("Enviando alterações para o GitHub...")
    subprocess.run(["git", "add", ARQUIVO_NO_REPO], cwd=CAMINHO_REPO_GIT, check=True)

    resultado = subprocess.run(
        ["git", "commit", "-m", MENSAGEM_COMMIT], cwd=CAMINHO_REPO_GIT
    )

    if resultado.returncode == 0:
        subprocess.run(["git", "push"], cwd=CAMINHO_REPO_GIT, check=True)
        print("Alterações enviadas com sucesso para o GitHub!")
    else:
        print("Nenhuma alteração nova para commitar (planilha sem mudanças).")


if __name__ == "__main__":
    atualizar_e_exportar()
    subir_para_github()
    print("\n============================================")
    print("Execução finalizada.")
    print("============================================")
    input("Pressione Enter para fechar esta janela...")
