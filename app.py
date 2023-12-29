from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options 
import time
import random
from datetime import datetime

import pyautogui as pg


# configurações iniciais da página
def iniciar_drive():

    chrome_options = Options()
    arguments = ['--lang=pt-US', '--window-size=800,600', '--incognito', '--disable-notifications']
    for argument in arguments:
        chrome_options.add_argument(argument)

    # configurações experimentais
    chrome_options.add_experimental_option('prefs', {
        # alterar local padrão de dowloads
        'dowload.default_directory': 'C:\\Users\\enriq\\Downloads',
        #notifiar o google sobre essa alteração
        'dowload.directory_upgrade': False,
        # Desabilitar a confirmação de dowload
        'dowload.prompt_for_dowload': False,
        # desabilitar notificações
        'profile.default_content_setting_values.notifications': 2,
        # permirir multiplos dowloads
        'profile.default_content_setting_values.automatic_dowloads': 1
    })

    # Iniciando o webdriver
    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=chrome_options)
    
    return driver


# funçao para simular um ser humano digitando
def digitar_naturalmente(texto, elemento):
    for letra in texto:
        elemento.send_keys(letra)
        time.sleep(random.randint(1, 5)/ 30)



# abrindo a página do meu email 
driver = iniciar_drive()
driver.get('https://outlook.com')


driver.maximize_window()
time.sleep(2)


# salvar a janela inicial
janela_inicial = driver.current_window_handle


# iniciar uma nova sessão, clicando em um botão para isso
iniciar_sessao = driver.find_element(By.CLASS_NAME, 'btn.btn-outlook' )
iniciar_sessao.click()
time.sleep(1)


# em outra janela
janelas = driver.window_handles
for janela in janelas:
    if janela not in janela_inicial:
        #alterar para essa nova janela
        driver.switch_to.window(janela)
        driver.maximize_window()
        time.sleep(1)

        # digitar o email
        try:
            email = driver.find_element('xpath', '*//input[@placeholder="E-mail, telefone ou Skype"]')
            email.click()
            time.sleep(1)

            email_texto = 'empresa.henrique@outlook.com'
            digitar_naturalmente(email_texto, email)

        except Exception as e:
            print(f'ERRO! não localizei o campo para digitar o seu email {e}')

        # salvar o email
        salvar_email = driver.find_element(By.XPATH, '*//input[@id="idSIButton9"]')
        salvar_email.click()
        time.sleep(1)

        try:
            # localizar e escrever a senha
            senha = driver.find_element(By.XPATH, '*//input[@id="i0118"]')
            senha.click()
            time.sleep(1)

            senha_texto = 'Rick260454*'
            digitar_naturalmente(senha_texto, senha)

        except Exception as e:
            print(f'ERRO! não encontrei o campo para digitar sua senha {e}')
        
        # salvar a senha
        salvar_senha = driver.find_element(By.XPATH, '*//input[@id="idSIButton9"]')
        salvar_senha.click()
        time.sleep(1)

        
        # clicar no checkbox (não mostrar novamente)
        checkbox = driver.find_element(By.XPATH, '*//input[@id="KmsiCheckboxField"]')
        checkbox.click()
        time.sleep(0.5)


        # Não Manter a sessão iniciada para que da próxima vez, o código não quebre
        btn_nao = driver.find_element(By.XPATH, '*//input[@id="idBtn_Back"]')
        btn_nao.click()
        time.sleep(2)


        # clicar em pesquisar e escrever 
        pesquisar = driver.find_element(By.XPATH, '*//input[@id="topSearchInput"]')
        pesquisar.click()
        time.sleep(1)

        pesq_texto = 'lpnet' 

        pesquisar.send_keys(pesq_texto)
        time.sleep(1)

        pesquisar.send_keys(Keys.ENTER)
        time.sleep(2)

        # filtrar por item que possuem anexos
        anexo = driver.find_element(By.XPATH, '*//button[@class="ms-Button ms-Button--default sK5B2 fp1Mc root-423"]')
        anexo.click()
        time.sleep(1)

        # clicar no email com o documento 
        email_com_anexo = driver.find_element(By.XPATH, '*//div[@class="hcptT gDC9O"]')
        email_com_anexo.click()
        time.sleep(3)

        # escolher a opção de dowload do arquivo
        opc = driver.find_element(By.XPATH, '*//button[@title="Mais ações"]')
        opc.click()
        time.sleep(2)

        # baixar o arquivo
        baixar = driver.find_element(By.XPATH, '*//button[@name="Baixar"]')
        baixar.click()
        time.sleep(1)



# clicar em dowloads
dowload = pg.click(116,371, duration=0.5)
time.sleep(0.05)


# alterar o none do arquivo
nome_arquivo = pg.click(312,474, duration=0.5)
time.sleep(0.5)
nome_arquivo = pg.press('backspace')
time.sleep(0.5)


# Obter a data atual
data_atual = datetime.now()

# Formatar a data no formato desejado (DD/MM)
data_formatada = data_atual.strftime('%d.%m')

# Criar a string completa
mensagem = f'Boleto do dia {data_formatada}'


pg.typewrite(mensagem)
time.sleep(0.5)

#salvar
salvar = pg.click(791,561, duration=0.5)
time.sleep(2)


# exluir o email para que um novo tome o seu lugar
excluir = pg.moveTo(785,481, duration=0.5)
excluir = pg.click()
time.sleep(1)


# confirmar a exclusão
btn_ok = pg.click(964,640, duration=0.5)
time.sleep(2)


# fechr as janelas
close = pg.click(1901,10, duration=0.5)
time.sleep(3)


# abrir o whatsapp
windows = pg.press('win')
time.sleep(1)
pg.typewrite('whatsapp')
time.sleep(1)
pg.press('enter')
time.sleep(5)


# pesquisar por pai
campo_pesquisa = pg.click(230,161, duration=0.5)
time.sleep(1)
campo_pesquisa = pg.typewrite('pai')
time.sleep(1)
campo_pesquisa = pg.press('tab')
time.sleep(1)
campo_pesquisa = pg.press('enter')
time.sleep(3)


# abrir o executar
pg.hotkey('win', 'r')
time.sleep(0.05)
pg.click(365,877, duration=0.5)

for i in range(30):
    pg.press('backspace')
    time.sleep(0.01)

pg.typewrite('C:\\Users\\enriq\\Downloads')
time.sleep(0.05)
pg.press('enter')
time.sleep(5)


# arrastar o último dowload até a mensagem e soltar
pg.moveTo(320,270, duration=0.5)
time.sleep(0.5)
pg.dragTo(1183,620, button='left', duration=0.5)
time.sleep(3)

# deixar uma legenda
lengenda = pg.click(744,896, duration=0.5)
time.sleep(1)


data_atual = datetime.now()
data_formatada = data_atual.strftime('%d/%m')
mensagem = f'Boleto do dia {data_formatada}, com data de vencimento para o dia 10 deste mês'

for letra in mensagem:
    pg.typewrite(letra)
    time.sleep(random.randint(1,5)/30)
time.sleep(1)

pg.press('enter')


driver.close()