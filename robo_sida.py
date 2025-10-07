import pandas as pd
import pyautogui as pa
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import time
import pyperclip

pa.PAUSE = 1

# --- CONFIGURAÇÃO DAS OPÇÕES DO CHROME ---
options = Options()
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

# IMPORTANDO A PLANILHA DO DIA

sida_df = pd.read_excel("sida_hoje.xlsx", header=None)
pd.set_option("display.max_rows", None)
print(sida_df)

primeira_vez = True

# ABRIR O NAVEGADOR
driver = webdriver.Chrome(options=options)
driver.maximize_window()
wait = WebDriverWait(driver, 15)

# LOGAR COM O EMAIL MANUALMENTE
driver.get(
    "https://drive.google.com/drive/folders/1JdMAtoFrlyWF3A7B-B7qbHiUUR_CnwW_")
time.sleep(45)

# ABRIR O SIDA
pa.hotkey("ctrl", "l")
pa.write("https://sida.pgfn.fazenda/sida/#/sida/consulta/busca")
pa.press("enter")
time.sleep(30)

# COMEÇO DO CADASTRAMENTO
for i in range(len(sida_df)):
    print(f"Processando a linha: {i}")

    # COPIA O PRIMEIRO NÚMERO DO PROCESSO
    valor_para_copiar = sida_df.loc[i, 0]
    pyperclip.copy(valor_para_copiar)

    # PESQUISAR O PROCESSO
    pa.press("tab", presses=5)
    pa.press("enter")
    pa.hotkey("ctrl", "v")
    pa.press("enter")

    # ENCONTRAR SE TEM UM OU MAIS SIDAS
    try:
        wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn2[ng-click*='selecionarTudo']"))).click()
        print(
            f"Múltiplos resultados encontrados na linha {i}. Clicando em 'Imprimir Inscrições'.")

        wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn2[ng-click=\"vm.abrePopUp('inscricoes')\"]"))).click()
    except TimeoutException:
        print(
            f"Apenas um resultado encontrado na linha {i}. Clicando em 'IMPRIMIR'.")
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[text()='IMPRIMIR']"))).click()

    # IMPRIMIR TELA
    wait.until(EC.visibility_of_element_located(
        (By.XPATH, "//button[text()='OK']")))
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//input[@value='R']"))).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[text()='OK']"))).click()

    # VOLTAR PARA A PAGINA INICIAL

    try:

        modal_bloqueador = (By.CSS_SELECTOR, "div.sida-modal")

        wait.until(EC.invisibility_of_element_located(modal_bloqueador))

        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[text()='VOLTAR']"))).click()
        print("Botão 'Voltar' encontrado e clicado.")

    except TimeoutException:

        wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn3[ng-click*='novaConsulta']"))).click()
        print("Botão 'Nova Consulta' encontrado. Voltando para a próxima busca.")

    time.sleep(2)

    # COPIAR O LINK DO DRIVE

    print("Fazendo o upload no drive")
    link_drive = sida_df.loc[i, 2]
    pyperclip.copy(link_drive)
    pa.hotkey("ctrl", "t")
    time.sleep(2)
    pa.hotkey("ctrl", "v")
    pa.press("enter")
    time.sleep(5)

    # UPLOAD DO SIDA NO DRIVE
    pa.hotkey("alt", "c")
    pa.press("u")
    time.sleep(7)

    if primeira_vez:
        time.sleep(10)
        pa.write("Downloads")
        pa.press("enter")
        time.sleep(3)
        primeira_vez = False

    pa.click(x=370, y=167)
    pa.press("enter")
    time.sleep(10)
    print("Upload concluído!")
    pa.hotkey("ctrl", "w")
    time.sleep(2)

print("Processamento de todas as linhas concluído!")
driver.quit()

