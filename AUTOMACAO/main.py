import webbrowser  # Para abrir o navegador
import pygetwindow as gw  # Para manipular as janelas abertas no sistema
from urllib.parse import quote
import os  # Para executar comandos no sistema operacional
import time  # Para controlar o tempo de espera
import pandas as pd  # Bibliotecas necessárias
import pyautogui  # Para automação de clique e digitação
from dotenv import load_dotenv

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

webbrowser.register("chrome", None, webbrowser.BackgroundBrowser(BROWSER_PATH))

# Carregar a planilha com os dados de WhatsApp
arquivo = CAMINHO_PLANILHA
planilha = pd.read_excel(arquivo)


# Função para enviar a mensagem
def enviar_mensagem(mensagem):
    # Esperar um tempo para garantir que o campo de digitação esteja visível
    print("Localizando o campo de digitação para a mensagem...")
    time.sleep(5)  # Aguardar 5 segundos para o campo de digitação estar visível

    try:
        # Localizar o campo de digitação da mensagem usando uma imagem do campo de texto
        campo_pos = pyautogui.locateOnScreen(
            CAMINHO_PNG, confidence=0.6
        )  # Ajuste a imagem conforme necessário

        if campo_pos:
            # Se o campo for encontrado, clicamos nele
            campo_centro = pyautogui.center(campo_pos)
            pyautogui.click(campo_centro)  # Clica no campo de texto
            print("Campo de mensagem selecionado.")

            pyautogui.press(
                "enter"
            )  # Simula pressionar a tecla Enter para enviar a mensagem
            print("Mensagem enviada!")
        else:
            print("Não foi possível localizar o campo de digitação.")
    except pyautogui.ImageNotFoundException as e:
        print("Não foi possível localizar a imagem.")
        raise e


# Iterar sobre os links de WhatsApp na planilha
for i, mensagem in enumerate(planilha["MENSAGEM"]):
    if pd.notnull(mensagem):  # Verifica se o link não está vazio
        mensagem_formatada = quote(mensagem)
        numero = planilha["NÚMERO"][i]
        link_whatsapp_web = (
            f"https://web.whatsapp.com/send?phone={numero}&text={mensagem_formatada}"
        )
        print(f"Abrindo o link {i + 1}: {link_whatsapp_web}")

        # Abrir o link no Opera GX (ou outro navegador)
        print("Abrindo o WhatsApp Web no navegador...")
        webbrowser.get("chrome").open(link_whatsapp_web)
        time.sleep(
            10
        )  # Espera a página carregar (pode ser ajustado conforme necessário)

        # Envia a mensagem da planilha para o whatsapp
        print(f"Enviando a mensagem: {mensagem}")
        enviar_mensagem(mensagem)

        # Fechar a aba atual do navegador após enviar a mensagem
        pyautogui.hotkey("ctrl", "w")  # Fecha a aba do navegador
        time.sleep(5)  # Pausa para garantir o fechamento da aba

# Mensagem final indicando que o processo foi concluído
print("Todos os links foram processados e as mensagens foram enviadas.")
