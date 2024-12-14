import webbrowser  # Para abrir o navegador
from urllib.parse import quote
import os  # Para executar comandos no sistema operacional
import time  # Para controlar o tempo de espera
import pandas as pd  # Bibliotecas necessárias
import pyautogui  # Para automação de clique e digitação
from dotenv import load_dotenv
from utils.browser import get_browser
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

load_dotenv()

CAMINHO_PLANILHA = os.environ.get(
    "CAMINHO_PLANILHA",
    None,  # Caminho do arquivo Excel definido no arquivo .env
)

CAMINHO_PNG = os.environ.get(
    "CAMINHO_PNG",
    None,
)

BROWSER_PATH = os.environ.get(
    "BROWSER_PATH",
    None,
)

PROFILE_CHROME_PATH = os.environ.get(
    "PROFILE_CHROME_PATH",
    None,
)

browser = get_browser(
    rf"user-data-dir={PROFILE_CHROME_PATH}",
    "--headless",
)

# Carregar a planilha com os dados de WhatsApp
arquivo = CAMINHO_PLANILHA
planilha = pd.read_excel(arquivo)


# Função para enviar a mensagem
def enviar_mensagem(mensagem):
    print("Localizando o campo de digitação para a mensagem...")

    try:
        # Localizar o campo de digitação da mensagem usando uma imagem do campo de texto
        inputMensagem = browser.find_element(
            By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]'
        )

        if inputMensagem:
            # Se o campo for encontrado, clicamos nele
            inputMensagem.click()
            time.sleep(1)
            actions = ActionChains(browser)
            actions.move_to_element(inputMensagem).click()
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(1)
        else:
            print("Não foi possível localizar o campo de digitação.")
    except pyautogui.ImageNotFoundException as e:
        print("Não foi possível localizar a imagem.")
        raise e


# Iterar sobre as mensagens de WhatsApp na planilha
for i, mensagem in enumerate(planilha["MENSAGEM"]):
    if pd.notnull(mensagem):  # Verifica se o campo de mensagem não está vazio
        mensagem_formatada = quote(mensagem)
        numero = planilha["NÚMERO"][i]
        link_whatsapp_web = (
            f"https://web.whatsapp.com/send?phone={numero}&text={mensagem_formatada}"
        )
        print(f"Abrindo o link {i + 1}: {link_whatsapp_web}")

        # Abrir o link no Opera GX (ou outro navegador)
        print("Abrindo o WhatsApp Web no navegador...")
        browser.get(link_whatsapp_web)

        # Envia a mensagem da planilha para o whatsapp
        print(f"Enviando a mensagem: {mensagem}")
        enviar_mensagem(mensagem)

browser.quit()  # Fechar o navegador
# Mensagem final indicando que o processo foi concluído
print("Todos os links foram processados e as mensagens foram enviadas.")
