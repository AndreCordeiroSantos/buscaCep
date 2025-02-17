import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyperclip
import os
import pyautogui
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def ler_ceps(caminho_arquivo):
    with open(caminho_arquivo, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)  # Ignora a primeira linha (cabeçalho)
        return [row[0] for row in reader]

def buscar_cep(cep, driver):
    driver.get("https://buscacep.com.br/")
    
    try:
        # Espera explícita para garantir que o elemento esteja presente
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
        search_box.send_keys(cep)
        search_box.send_keys(Keys.RETURN)
        
        # Espera 3 segundos após inserir o CEP
        time.sleep(3)
        
        # Espera explícita para garantir que o resultado esteja presente
        resultado = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".table"))
        )
        
        # Extrai os dados da tabela
        dados = {}
        try:
            dados['CEP'] = driver.find_element(By.ID, 'cepTd').text
            dados['Logradouro'] = driver.find_element(By.ID, 'logradouroTd').text
            dados['Bairro'] = driver.find_element(By.ID, 'bairroTd').text
            dados['Localidade'] = driver.find_element(By.ID, 'localidadeTd').text
            dados['UF'] = driver.find_element(By.ID, 'ufTd').text
            dados['Ibge'] = driver.find_element(By.ID, 'ibgeTd').text
            dados['DDD'] = driver.find_element(By.ID, 'dddTd').text
        except:
            dados['CEP'] = cep
            dados['Logradouro'] = 'CEP não encontrado'
            dados['Bairro'] = 'CEP não encontrado'
            dados['Localidade'] = 'CEP não encontrado'
            dados['UF'] = 'CEP não encontrado'
            dados['Ibge'] = 'CEP não encontrado'
            dados['DDD'] = 'CEP não encontrado'
        
        return dados
        
    except Exception as e:
        print(f"Erro ao buscar o CEP {cep}: {e}")
        dados = {
            'CEP': cep,
            'Logradouro': 'CEP não encontrado',
            'Bairro': 'CEP não encontrado',
            'Localidade': 'CEP não encontrado',
            'UF': 'CEP não encontrado',
            'Ibge': 'CEP não encontrado',
            'DDD': 'CEP não encontrado'
        }
        return dados

def salvar_resultados(resultados, caminho_arquivo):
    with open(caminho_arquivo, mode='w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['CEP', 'Logradouro', 'Bairro', 'Localidade', 'UF', 'Ibge', 'DDD']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for resultado in resultados:
            writer.writerow(resultado)

def gerar_pdf(resultados, caminho_pdf):
    c = canvas.Canvas(caminho_pdf, pagesize=letter)
    width, height = letter
    
    def draw_header_footer(c, page_num):
        # Cabeçalho
        c.setFont("Helvetica-Bold", 16)
        c.drawString(30, height - 40, "Relatório de Resultados de CEP")
        
        # Rodapé
        c.setFont("Helvetica", 10)
        c.drawString(30, 30, f"Criado por André Rafael @2025 - Página {page_num}")
    
    y = height - 80
    page_num = 1
    draw_header_footer(c, page_num)
    
    c.setFont("Helvetica", 12)
    for resultado in resultados:
        for key, value in resultado.items():
            c.drawString(30, y, f"{key}: {value}")
            y -= 20
            if y < 50:
                c.showPage()
                page_num += 1
                draw_header_footer(c, page_num)
                c.setFont("Helvetica", 12)
                y = height - 80
        y -= 10  # Espaço entre registros
    
    c.save()

# def login_gmail(driver, email, password):
#     driver.get("https://gmail.com")
    
#     try:
#         # Espera explícita para garantir que o campo de email esteja presente
#         email_field = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.ID, "identifierId"))
#         )
#         email_field.send_keys(email)
#         email_field.send_keys(Keys.RETURN)
        
#         # Espera explícita para garantir que o campo de senha esteja presente
#         password_field = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.NAME, "Passwd"))
#         )
#         password_field.send_keys(password)
#         password_field.send_keys(Keys.RETURN)
        
    # except Exception as e:
    #     print(f"Erro ao fazer login no Gmail: {e}")

if __name__ == "__main__":
    base_path = os.path.dirname(__file__)
    caminho_arquivo = os.path.join(base_path, 'cep-list.csv')
    ceps = ler_ceps(caminho_arquivo)
    
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    resultados = []
    try:
        for cep in ceps:
            resultado = buscar_cep(cep, driver)
            if resultado:
                resultados.append(resultado)
    finally:
        driver.quit()
    
    caminho_resultados = os.path.join(base_path, 'resultados.csv')
    salvar_resultados(resultados, caminho_resultados)
    
    caminho_pdf = os.path.join(base_path, 'relatorio_resultados.pdf')
    gerar_pdf(resultados, caminho_pdf)
    
    ## Fazer login no Gmail via PYAUTOGUI
    def abrir_edge():
        # Abrir o Executar
        pyautogui.hotkey('win', 'r')
        time.sleep(1)
        
        # Digitar o comando para abrir o Microsoft Edge
        pyautogui.typewrite('msedge')
        pyautogui.press('enter')
        time.sleep(3)  # Espera o Edge abrir

    def acessar_gmail():
        # Posicionar o mouse na barra de URL (ajuste as coordenadas conforme necessário)
        pyautogui.click(x=300, y=50)
        time.sleep(1)
        
        # Digitar o URL do Gmail e pressionar Enter
        pyautogui.typewrite('gmail.com')
        pyautogui.press('enter')
        time.sleep(5)  # Espera a página carregar

    def login_gmail(email, password):
        try:
            # Digitar o email
            pyautogui.typewrite(email)
            pyautogui.press('enter')
            time.sleep(3)  # Espera a próxima página carregar
            
            # Digitar a senha
            pyautogui.typewrite(password)
            pyautogui.press('enter')
            time.sleep(5)  # Espera o login ser concluído
        except Exception as e:
            print(f"Erro ao fazer login no Gmail: {e}")

    def enviar_email(destinatario, assunto, caminho_anexo):
        try:
            # Clicar no botão "Escrever"
            pyautogui.click(x=100, y=200)  # Ajuste as coordenadas conforme necessário
            time.sleep(2)
            
            # Digitar o destinatário
            pyautogui.typewrite(destinatario)
            pyautogui.press('tab')
            time.sleep(1)
            
            # Digitar o assunto
            pyautogui.typewrite(assunto)
            pyautogui.press('tab')
            time.sleep(1)
            
            # Digitar o corpo do email
            corpo_email = """Olá,

            Aqui estão os resultados do CEP solicitado.
            Atenciosamente,
            Sua Empresa

            Email criado e feito por André Cordeiro."""
            pyautogui.typewrite(corpo_email)
            time.sleep(1)
            
            # Dar 3 tabs para chegar no ícone de anexar arquivo
            pyautogui.press('tab', presses=3, interval=0.5)
            pyautogui.press('enter')
            time.sleep(6)  # Espera a janela de anexar arquivo abrir
            
            # Copiar o caminho do arquivo para a área de transferência e colar
            pyperclip.copy(caminho_anexo)
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('enter')
            time.sleep(3)  # Espera o arquivo ser anexado
            
            # Dar 1 tab para chegar no botão "Enviar"
            pyautogui.press('tab')
            pyautogui.press('enter')
            time.sleep(2)
            
        except Exception as e:
            print(f"Erro ao enviar o email: {e}")

    if __name__ == "__main__":
        # Abrir o Microsoft Edge usando pyautogui
        abrir_edge()
        
        # Acessar o Gmail
        acessar_gmail()
        
        # Fazer login no Gmail
        login_gmail("*exemploemailDeEnvio@gmail.com", "*senha email para login")
        
        # Enviar o email com o anexo
        enviar_email("*exemplodeemail@gmail.com", "Assunto email", "#caminho do arquivo")