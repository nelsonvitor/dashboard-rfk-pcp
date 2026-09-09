<<<<<<< HEAD
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
import traceback
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

    # Evita o pop-up "Este arquivo contém vínculos para outros arquivos.
    # Deseja atualizá-los?" que trava o script esperando clique manual.
    excel.AskToUpdateLinks = False

    try:
        print("Abrindo planilha de origem...")
        # UpdateLinks=0 -> não atualiza vínculos externos ao abrir (evita o prompt acima)
        wb = excel.Workbooks.Open(CAMINHO_PLANILHA_ORIGEM, UpdateLinks=0)

        # Se o arquivo foi copiado/baixado/sincronizado (OneDrive, rede, e-mail etc.),
        # o Windows pode marcá-lo como "bloqueado" e o Excel abre em Modo de Exibição
        # Protegida (somente leitura), exigindo clique manual em "Habilitar Edição".
        # O bloco abaixo detecta isso e libera a edição automaticamente.
        if excel.ProtectedViewWindows.Count > 0:
            print("Arquivo abriu em Modo de Exibição Protegida. Habilitando edição automaticamente...")
            pvw = excel.ProtectedViewWindows.Item(1)
            wb = pvw.Edit()  # converte a janela protegida em um Workbook editável normal

        if wb.ReadOnly:
            raise RuntimeError(
                "A planilha foi aberta como SOMENTE LEITURA e não foi possível liberar "
                "a edição automaticamente. Verifique se o arquivo não está aberto por "
                "outra pessoa/processo, ou se está marcado como 'Somente leitura' nas "
                "propriedades do arquivo no Windows (botão direito > Propriedades > "
                "desmarcar 'Somente leitura')."
            )

        print("Atualizando Power Query (RefreshAll)...")
        wb.RefreshAll()

        # Consultas do Power Query rodam em segundo plano (assíncronas).
        # Isso força o Excel a esperar até todas terminarem.
        excel.CalculateUntilAsyncQueriesDone()

        # Pequena margem de segurança extra
        time.sleep(3)

        print("Salvando planilha de origem com os dados atualizados...")
        wb.Save()

        print(f"Exportando aba '{NOME_ABA_TABELA}' como CSV...")
        aba = wb.Worksheets(NOME_ABA_TABELA)
        aba.Copy()  # cria um novo workbook temporário só com essa aba
        novo_wb = excel.ActiveWorkbook
        novo_wb.SaveAs(CAMINHO_SAIDA_CSV, FileFormat=62)  # 62 = CSV UTF-8 (preserva acentos como em "MÊS")
        novo_wb.Close(SaveChanges=False)

        wb.Close(SaveChanges=False)  # já foi salva explicitamente acima com wb.Save()
        print(f"CSV salvo com sucesso em: {CAMINHO_SAIDA_CSV}")
    finally:
        excel.Quit()


def subir_para_github():
    print("Enviando alterações para o GitHub...")
    subprocess.run(["git", "add", ARQUIVO_NO_REPO], cwd=CAMINHO_REPO_GIT, check=True)

    agora = time.strftime("%d/%m/%Y %H:%M:%S")

    # Verifica se há mudanças reais no conteúdo do arquivo
    resultado_status = subprocess.run(
        ["git", "diff", "--cached", "--quiet", ARQUIVO_NO_REPO],
        cwd=CAMINHO_REPO_GIT,
    )
    houve_mudanca = resultado_status.returncode != 0

    if houve_mudanca:
        mensagem = f"{MENSAGEM_COMMIT} - {agora}"
        subprocess.run(["git", "commit", "-m", mensagem], cwd=CAMINHO_REPO_GIT, check=True)
    else:
        # Cria um commit vazio só para registrar que o script rodou e verificou os dados
        mensagem = f"Verificação automática (sem mudanças nos dados) - {agora}"
        subprocess.run(
            ["git", "commit", "--allow-empty", "-m", mensagem],
            cwd=CAMINHO_REPO_GIT,
            check=True,
        )

    subprocess.run(["git", "push"], cwd=CAMINHO_REPO_GIT, check=True)
    print("Alterações enviadas com sucesso para o GitHub!")


if __name__ == "__main__":
    try:
        atualizar_e_exportar()
        subir_para_github()
        print("\n============================================")
        print("Execução finalizada.")
        print("============================================")
    except Exception:
        print("\n============================================")
        print("ERRO durante a execução do script:")
        print("============================================")
        traceback.print_exc()
        print("============================================")
    finally:
        # Isso garante que a janela SEMPRE fique aberta ao final,
        # mesmo que tenha ocorrido um erro acima.
        input("Pressione Enter para fechar esta janela...")
=======
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
>>>>>>> d28ccc7dd4302997f01ff5e5ac7296f4433e62cc
