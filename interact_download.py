"""
interact_download.py
====================
Baixa os dois relatórios do Interact (Dados e Unidade) em modo headless
e salva no repositório para os scripts de atualização usarem.

Roda via GitHub Actions todo dia às 6h (BRT).
"""

import time
import os
import sys
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ── CONFIGURAÇÕES ──────────────────────────────────────────────────────────────
LOGIN = os.environ["INTERACT_LOGIN"]        # Secret no GitHub
SENHA = os.environ["INTERACT_SENHA"]        # Secret no GitHub

DOWNLOAD_DIR = str(Path(__file__).parent / "downloads")
Path(DOWNLOAD_DIR).mkdir(exist_ok=True)

RELATORIOS = [
    {
        "nome": "RTA Dados",
        "url": "https://clealco.interact.com.br/sa/go.jsp?to=_bll4cld1MTZGV2taVWE1Z2xRVFM3UG1kVnl4WEh1MC9UV2ZkS2t3eTN4TEZDakd4dVlpdkF4R3cyOVVlb0NJZnBOenE2T3E4Z25BcHM1USsybEI2cVE9PQ==",
        "arquivo": "RTA_Dados.xlsx",
    },
    {
        "nome": "RTA Unidade",
        "url": "https://clealco.interact.com.br/sa/go.jsp?to=_bll4cld1MTZGV2taVWE1Z2xRVFM3UG1kVnl4WEh1MC9UV2ZkS2t3eTN4TCtJU3AvZGs3eCt4V2hZdW9xeXlKVnBOenE2T3E4Z25BcHM1USsybEI2cVE9PQ==",
        "arquivo": "RTA_Unidade.xlsx",
    },
]


# ── HELPERS ────────────────────────────────────────────────────────────────────
def snapshot(pasta: str) -> set:
    return set(os.listdir(pasta))


def aguardar_download(pasta: str, antes: set, timeout: int = 240) -> str:
    limite = time.time() + timeout
    while time.time() < limite:
        agora = set(os.listdir(pasta))
        novos = [f for f in (agora - antes) if not f.endswith(".crdownload")]
        if novos:
            paths = [os.path.join(pasta, f) for f in novos if os.path.exists(os.path.join(pasta, f))]
            if paths:
                return max(paths, key=os.path.getctime)
        time.sleep(1)
    raise TimeoutError("Download não completou no tempo limite.")


def localizar_login_senha(driver):
    driver.switch_to.default_content()

    def buscar():
        senha = driver.find_element(By.XPATH, "//input[@type='password']")
        login = senha.find_element(By.XPATH, "preceding::input[@type='text'][1]")
        return login, senha

    try:
        return buscar()
    except Exception:
        pass

    for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            driver.switch_to.frame(iframe)
            return buscar()
        except Exception:
            driver.switch_to.default_content()

    return None, None


def converter_xls_para_xlsx(caminho_xls: str) -> str:
    """Converte .xls para .xlsx usando openpyxl (funciona no Linux/GitHub Actions)."""
    import xlrd
    import openpyxl

    caminho_xlsx = os.path.splitext(caminho_xls)[0] + ".xlsx"

    wb_xls = xlrd.open_workbook(caminho_xls)
    wb_xlsx = openpyxl.Workbook()

    for idx, sheet_name in enumerate(wb_xls.sheet_names()):
        sheet_xls = wb_xls.sheet_by_name(sheet_name)
        if idx == 0:
            ws = wb_xlsx.active
            ws.title = sheet_name
        else:
            ws = wb_xlsx.create_sheet(title=sheet_name)

        for row in range(sheet_xls.nrows):
            ws.append(sheet_xls.row_values(row))

    wb_xlsx.save(caminho_xlsx)
    os.remove(caminho_xls)
    return caminho_xlsx


# ── CHROME HEADLESS ────────────────────────────────────────────────────────────
def criar_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    return webdriver.Chrome(options=options)


# ── DOWNLOAD DE UM RELATÓRIO ───────────────────────────────────────────────────
def ja_logado(driver) -> bool:
    """Verifica se já está autenticado (sem campos de login visíveis)."""
    try:
        driver.find_element(By.XPATH, "//input[@type='password']")
        return False  # Achou campo de senha = não está logado
    except Exception:
        return True  # Sem campo de senha = já logado


def baixar_relatorio(driver, wait, relatorio: dict) -> str:
    print(f"\n{'='*50}")
    print(f"📥 Baixando: {relatorio['nome']}")

    driver.get(relatorio["url"])
    time.sleep(8)

    # Login — só faz se ainda não estiver logado
    if not ja_logado(driver):
        login_input, senha_input = localizar_login_senha(driver)
        if not login_input or not senha_input:
            raise Exception("Não foi possível localizar campos de login.")

        driver.execute_script(
            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
            login_input, LOGIN
        )
        driver.execute_script(
            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
            senha_input, SENHA
        )
        senha_input.send_keys(Keys.ENTER)
        print("✅ Login realizado")
        time.sleep(12)
    else:
        print("✅ Sessão já ativa — sem necessidade de novo login")
        time.sleep(5)

    # Exportar
    icone_exportar = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//*[contains(@class,'export') or contains(@title,'Export')]"
        ))
    )

    antes = snapshot(DOWNLOAD_DIR)
    driver.execute_script("arguments[0].click();", icone_exportar)
    print("🟢 Ícone Exportar clicado")
    time.sleep(2)

    opcao_planilha = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//*[contains(text(),'Planilha') or contains(text(),'Excel')]"
        ))
    )
    driver.execute_script("arguments[0].click();", opcao_planilha)
    print("🟢 Exportação iniciada")

    # Aguardar download
    print("⏳ Aguardando download...")
    arquivo = aguardar_download(DOWNLOAD_DIR, antes)
    print(f"⬇️ Arquivo baixado: {os.path.basename(arquivo)}")

    # Converter se necessário
    if arquivo.lower().endswith(".xls"):
        print("🔄 Convertendo .xls → .xlsx...")
        arquivo = converter_xls_para_xlsx(arquivo)
        print(f"✅ Convertido: {os.path.basename(arquivo)}")

    # Renomear para nome fixo
    destino = os.path.join(DOWNLOAD_DIR, relatorio["arquivo"])
    if os.path.exists(destino):
        os.remove(destino)
    os.rename(arquivo, destino)
    print(f"✅ Salvo como: {relatorio['arquivo']}")

    return destino


# ── MAIN ───────────────────────────────────────────────────────────────────────
def main():
    driver = criar_driver()
    wait = WebDriverWait(driver, 40)
    erros = []

    try:
        for relatorio in RELATORIOS:
            try:
                baixar_relatorio(driver, wait, relatorio)
            except Exception as e:
                print(f"❌ Erro em {relatorio['nome']}: {e}")
                erros.append(relatorio['nome'])
    finally:
        driver.quit()

    if erros:
        print(f"\n❌ Falhas: {', '.join(erros)}")
        sys.exit(1)
    else:
        print("\n✅ Todos os relatórios baixados com sucesso!")


if __name__ == "__main__":
    main()
