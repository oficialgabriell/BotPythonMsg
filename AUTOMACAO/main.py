from urllib.parse import quote  # Para codificar a mensagem
import os  # Para executar comandos no sistema operacional
import time  # Para controlar o tempo de espera
import pandas as pd  # Bibliotecas necessárias para manipular planilhas Excel
from dotenv import load_dotenv  # Para carregar variáveis de ambiente
from utils.browser import get_browser  # Função para obter o navegador
from selenium.webdriver.common.by import By  # Para localizar elementos na página
from selenium.webdriver.common.keys import Keys  # Para simular teclas do teclado
from selenium.webdriver.common.action_chains import (
    ActionChains,
)  # Para simular ações do mouse
from selenium.common.exceptions import NoSuchElementException  # Para tratar exce


load_dotenv()


def get_env_variable(nome_variavel):
    value = os.environ.get(nome_variavel)
    if value is None:
        raise EnvironmentError(
            f"A variável de ambiente '{nome_variavel}' não está definida. Verifique o arquivo .env."
        )
    return value


CAMINHO_PLANILHA = get_env_variable("CAMINHO_PLANILHA")
BROWSER_PATH = get_env_variable("BROWSER_PATH")
PROFILE_CHROME_PATH = get_env_variable("PROFILE_CHROME_PATH")

BROWSER = get_browser(
    rf"user-data-dir={PROFILE_CHROME_PATH}", # Caminho do perfil do Chrome
    # "--headless",
)

# Carregar a planilha com os dados de WhatsApp
ARQUIVO = CAMINHO_PLANILHA
PLANILHA = pd.read_excel(ARQUIVO)


# Função para enviar a mensagem
def enviar_mensagem():
    print("Localizando o campo de digitação para a mensagem...")

    try:
        # Localizar o campo de digitação da mensagem usando uma imagem do campo de texto
        input_mensagem = BROWSER.find_element(
            By.XPATH,
            '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]',
        )
        input_mensagem.click()
        time.sleep(1)
        actions = ActionChains(BROWSER)
        actions.move_to_element(input_mensagem).click()
        actions.send_keys(Keys.ENTER)
        actions.perform()
        time.sleep(1)
    except NoSuchElementException:
        print("Não foi possível localizar o input de mensagem.")


def main(planilha):
    # Iterar sobre as mensagens de WhatsApp na planilha
    for i, mensagem in enumerate(planilha["MENSAGEM"]):
        if pd.notnull(mensagem):  # Verifica se o campo de mensagem não está vazio
            mensagem_formatada = quote(mensagem)
            numero = planilha["NÚMERO"][i]
            link_whatsapp_web = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem_formatada}"
            print(f"Abrindo o link {i + 1}: {link_whatsapp_web}")

            # Abrir o link no Opera GX (ou outro navegador)
            print("Abrindo o WhatsApp Web no navegador...")
            BROWSER.get(link_whatsapp_web)

            # Envia a mensagem da planilha para o whatsapp
            print(f"Enviando a mensagem: {mensagem}")
            enviar_mensagem()

    BROWSER.quit()  # Fechar o navegador
    # Mensagem final indicando que o processo foi concluído
    print("Todos os links foram processados e as mensagens foram enviadas.")


if __name__ == "__main__":
    main(PLANILHA)
