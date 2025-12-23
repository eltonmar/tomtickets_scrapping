
import os
from pathlib import Path
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
load_dotenv(ENV_PATH)


DOWNLOAD_PATH = r"C:\Users\BG-PROVISORIO\Desktop\Tickets do Tom"  # Caminho para salvar
Path(DOWNLOAD_PATH).mkdir(parents=True, exist_ok=True)  #nao sei o que isso faz nao - foi o chat que indicou


EMAIL = os.getenv("EMAIL_TOM")
PASSWORD = os.getenv("SENHA_TOM")
CONTA = os.getenv("CONTA_TOM")


def wait_new_download(download_dir: str, before_files: set, timeout: int = 120) -> Path:

    download_dir = Path(download_dir)
    end = time.time() + timeout

    while time.time() < end:
        cr = list(download_dir.glob("*.crdownload"))
        current_files = set(p.name for p in download_dir.glob("*") if p.is_file())

        new_files = current_files - before_files
        # remover essa parte do codigo ajuda pra melhorar a performance, pq eu ja sei que eh um xlms (Fonte: Gemini)
        candidates = [download_dir / name for name in new_files if not name.endswith(".crdownload")]

        if candidates and not cr:
            newest = max(candidates, key=lambda p: p.stat().st_mtime)
            return newest

        time.sleep(1)

    raise TimeoutError("Timeout esperando o download finalizar.")


def click_execute_by_report_name(driver, wait, report_name: str, download_dir: str, timeout_download: int = 120) -> Path:
    # espera a tabela carregar
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))


    before_files = set(p.name for p in Path(download_dir).glob("*") if p.is_file())

    row_xpath = f"//tr[.//td[normalize-space()='{report_name}'] or .//*[normalize-space()='{report_name}']]"
    row = wait.until(EC.presence_of_element_located((By.XPATH, row_xpath)))

    exec_link = row.find_element(By.XPATH, ".//td[contains(@class,'print-column')]//a[@title='Executar']")

    # como abre nova aba (target=_blank), salvamos a aba atual
    main_window = driver.current_window_handle
    existing_handles = set(driver.window_handles)

    driver.execute_script("arguments[0].click();", exec_link)

    end = time.time() + 15
    new_handle = None
    while time.time() < end:
        handles_now = set(driver.window_handles)
        diff = handles_now - existing_handles
        if diff:
            new_handle = list(diff)[0]
            break
        time.sleep(0.2)

    # se abriu nova aba, troca pra ela
    if new_handle:
        driver.switch_to.window(new_handle)

    # espera download terminar
    downloaded_file = wait_new_download(download_dir, before_files, timeout=timeout_download)

    # fecha aba extra e volta pra principal
    if new_handle:
        driver.close()
        driver.switch_to.window(main_window)

    return downloaded_file

#so precisa disso pra abrir o chrome
options = Options()

prefs = {
    "download.default_directory": DOWNLOAD_PATH,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options) #funcao chamada Options eh pica
wait = WebDriverWait(driver, 20)


try:
    driver.get("https://console.tomticket.com/login")

    email_input = wait.until(
        EC.presence_of_element_located((By.NAME, "conta"))
    )
    email_input.send_keys(CONTA)
    email_input = wait.until(
        EC.presence_of_element_located((By.NAME, "email"))
    )
    email_input.send_keys(EMAIL)

    senha_input = wait.until(
        EC.presence_of_element_located((By.NAME, "senha"))
    )
    senha_input.clear()
    senha_input.send_keys(PASSWORD)

    senha_input.send_keys(Keys.ENTER)

    #So pra esperar ra ver se o login vai funfar
    wait.until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    print("Login realizado com sucesso!")

    time.sleep(2)

    #Agora precisamos ir para a parte que eu vou acessar os Reports

    #Primeiro Report
    url_page1 = "https://console.tomticket.com/dashboard/reports/custom-reports"
    driver.get(url_page1)

    file1 = click_execute_by_report_name(driver, wait, "0007 - teste", DOWNLOAD_PATH, timeout_download=180)
    print(f"Baixado (página 1): {file1}")

    #Segundo report
    url_page3 = "https://console.tomticket.com/dashboard/reports/custom-reports?orderColumn=nome&direction=up&page=3"
    driver.get(url_page3)

    file2 = click_execute_by_report_name(driver, wait, "xx", DOWNLOAD_PATH, timeout_download=180)
    print(f"Baixado (página 3): {file2}")

    #url_reports = "https://console.tomticket.com/dashboard/reports/custom-reports"

finally:
    driver.quit()
