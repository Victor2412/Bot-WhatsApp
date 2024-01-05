import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


def fercharAba():
     pyautogui.hotkey("ctrl", "w")


webbrowser.open("https://web.whatsapp.com/")
sleep(30)

workbook = openpyxl.load_workbook("clientes.xlsx")
pagina_clientes = workbook["Sheet1"]

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    data = vencimento.strftime("%d/%m/%Y")

    mensagem = f"Ola {nome} seu boleto vence dia {data}. Por favor pagar no link https://www.link_pagamento.com"

    try:
        link_mensagem_whatsapp = (
            f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
        )
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen("seta_branca.png")
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        fercharAba()
        sleep(5)
    except:
        print(f"Não foi possível enviar mensagem para {nome}")
        with open("erros.csv", "a", newline="", encoding="utf-8") as arquivo:
            arquivo.write(f" {nome},{telefone}")
