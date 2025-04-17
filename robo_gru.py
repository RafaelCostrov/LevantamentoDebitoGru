import datetime
import os
import time
import base64
from dotenv import load_dotenv
from servico_google import acessando_sheets
from envio_drive import salvar_drive
from envio_email import enviar
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.print_page_options import PrintOptions
from selenium.common.exceptions import NoSuchElementException

load_dotenv()
client = acessando_sheets()
SHEET_LEVANTAMENTO_DEBITO = os.getenv('SHEET_LEVANTAMENTO_DEBITO')
spreadsheet = client.open_by_key(SHEET_LEVANTAMENTO_DEBITO)
guarulhos = spreadsheet.worksheet("Guarulhos")

coluna_nome = guarulhos.col_values(2)[1:]
coluna_resp = guarulhos.col_values(3)[1:]
coluna_cnpj = guarulhos.col_values(4)[1:]

chrome_options = Options()
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--headless")
chrome_options.add_experimental_option(
    "excludeSwitches", ["enable-automation"])
SERVICE = os.getenv('SERVICE')
servico = Service(SERVICE)
navegador = webdriver.Chrome(service=servico, options=chrome_options)

navegador.switch_to.window(navegador.window_handles[0])
navegador.get("https://fazenda.guarulhos.sp.gov.br/ords/guarulho/f?p=628:9")
navegador.maximize_window()

responsaveis_com_pasta = set()

for i, (cnpj, resp, nome) in enumerate(zip(coluna_cnpj, coluna_resp, coluna_nome), start=2):
    while True:
        try:
            navegador.get(
                "https://fazenda.guarulhos.sp.gov.br/ords/guarulho/f?p=628:9")
            navegador.maximize_window()
            extrato_debito = navegador.find_element(
                By.XPATH, '//*[@id="P9_EXTRATO_DEBITO"]/div/div[2]/div/a')
            extrato_debito.click()

            tipo_inscricao = navegador.find_element(
                By.XPATH, '//*[@id="P101_TIPO_INSCRICAO"]')
            tipo_inscricao.click()
            tipo_inscricao.send_keys("m", Keys.ENTER)

            campo_cnpj = navegador.find_element(
                By.XPATH, '//*[@id="P101_PESQUISA"]')
            campo_cnpj.clear()
            campo_cnpj.send_keys(cnpj)

            botao_pesquisar = navegador.find_element(
                By.XPATH, '//*[@id="P101_BTN_PESQUISA"]')
            botao_pesquisar.click()
            navegador.implicitly_wait(2)

            try:
                numero_inscricao = navegador.find_element(
                    By.XPATH, '//*[@id="report_P101_REL_PESQUISA"]/tbody/tr[2]/td/table/tbody/tr[2]/td[1]')
                numero_inscricao.click()
            except Exception as e:
                botao_home = navegador.find_element(
                    By.XPATH, '//*[@id="container"]/div[1]/div[1]/div[1]/a/div')
                botao_home.click()
                print(f"{cnpj}: sem inscricao municipal")
                break

            WebDriverWait(navegador, 30).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="container"]/div[2]/table/tbody/tr/td[1]/div/div/table/tbody/tr[5]/td/table/tbody/tr[1]/td'))
            )

            totais_element = navegador.find_element(
                By.XPATH, "//td[contains(text(), 'Totais :')]//following-sibling::td")
            valor_totais = totais_element.text
            print(valor_totais)

            if valor_totais != "":
                try:

                    dia = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                    nome_arquivo = f"{nome} - {cnpj} - {dia}.pdf"

                    print_options = PrintOptions()
                    print_options.shrink_to_fit = True
                    print_options.orientation = 'portrait'
                    pdf_content = navegador.print_page(print_options)

                    try:
                        pdf_bytes = base64.b64decode(pdf_content)
                    except Exception as e:
                        raise ValueError(f"Erro ao decodificar Base64: {e}")

                    try:
                        with open(nome_arquivo, "wb") as f:
                            f.write(pdf_bytes)
                    except Exception as e:
                        raise IOError(f"Erro ao salvar o arquivo PDF: {e}")

                except Exception as e:
                    print(f"Erro inesperado: {e}")

                if os.path.exists(nome_arquivo):
                    pasta_id = salvar_drive(nome_arquivo, resp)
                    responsaveis_com_pasta.add((resp, pasta_id))
                    time.sleep(5)
                    os.remove(nome_arquivo)
                else:
                    print("Erro: O arquivo PDF não foi gerado corretamente.")

            else:
                print("Empresa com valor de debito zerado: " + nome)

            botao_home = navegador.find_element(
                By.XPATH, '//*[@id="container"]/div[1]/div[1]/div[1]/a/div')
            botao_home.click()
            botao_sair = navegador.find_element(
                By.XPATH, '//*[@id="P9_DADOS"]/div/table/tbody/tr[1]/td[3]/a')
            botao_sair.click()
            break

        except Exception as e:
            print(f"Erro ao processar {cnpj}: {e}")
            navegador.close()
            time.sleep(1)
            navegador.quit()
            time.sleep(3)
            navegador = webdriver.Chrome(
                service=servico, options=chrome_options)

    print(responsaveis_com_pasta)

AMANDA = os.getenv('AMANDA')
AMANDA_O = os.getenv('AMANDA_O')
DANIELA_VIVIANE = os.getenv('DANIELA_VIVIANE')
LENI = os.getenv('LENI')
MARCIA = os.getenv('MARCIA')
TATIANE = os.getenv('TATIANE')
DEFAULT = os.getenv('DEFAULT')

for responsavel, id in responsaveis_com_pasta:
    link_pasta = f'https://drive.google.com/drive/folders/{id}'
    mes_atual = datetime.datetime.now().strftime("%m/%Y")

    def obter_email(responsavel):
        match responsavel:
            case "AMANDA":
                return AMANDA
            case "AMANDA O.":
                return AMANDA_O
            case "DANIELA VIVIANE":
                return DANIELA_VIVIANE
            case "LENI":
                return LENI
            case "MARCIA":
                return MARCIA
            case "TATIANE":
                return TATIANE
            case _:
                return DEFAULT
    enviar(obter_email(responsavel),
           f'Levantamento de Debito - {mes_atual}', responsavel.title(), link_pasta)

navegador.close()
time.sleep(1)
navegador.quit()
