import pyautogui
import time
import pandas as pd
import pyperclip

# Tempo para o computador Processar as Informações #
pyautogui.alert('O código vai começar, não mexa em nada enquanto ele estiver em execução.')
pyautogui.PAUSE = 1

# Abrir Menu Windows e Abrir o Firefox #
pyautogui.press('win')
pyautogui.write('firefox')
pyautogui.press('enter')

# Abrir a Caixa de Entrada do Email #
pyautogui.write('https://mail.google.com/mail/u/0/?hl=pt_BR#inbox')
pyautogui.press('enter')

#Verificar se o Site abriu corretamente #
while not pyautogui.locateOnScreen(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + rpa\icones_rpa\gmail_logo.png'):
    time.sleep(1)

# Localizar Contatos e Entrar na Pagina #
x, y, largura, altura = pyautogui.locateOnScreen(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + rpa\icones_rpa\menu_gmail.png')
pyautogui.click(x + largura/2, y + altura/2)
time.sleep(1)

x, y, largura, altura = pyautogui.locateOnScreen(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + rpa\icones_rpa\contatos_gmail.png')
pyautogui.click(x + largura/2, y + altura/2)

# Verificar se abriu a pagina contatos #
while not pyautogui.locateOnScreen(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + rpa\icones_rpa\pag_contatos.png'):
    time.sleep(1)

# exportar contatos #
x, y, largura, altura = pyautogui.locateOnScreen(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + rpa\icones_rpa\exportar_contatos.png')
pyautogui.click(x + largura/2, y + altura/2)

x, y, largura, altura = pyautogui.locateOnScreen(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + rpa\icones_rpa\confirmar_exportar.png')
pyautogui.click(x + largura/2, y + altura/2)

# Escrever email #

time.sleep(2)
contatos_df = pd.read_csv(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + rpa\contacts.csv', sep=';')
contatos_df = contatos_df.dropna(axis=1)

pyautogui.hotkey('ctrl', 'pgup')

for email in contatos_df['E-mail 1 - Value']:
    # Enviar email para todos da coluna #
    time.sleep(1)
    x, y, largura, altura = pyautogui.locateOnScreen(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + rpa\icones_rpa\escrever_gmail.png')
    pyautogui.click(x + largura/2, y + altura/2)

    # Escrever endereço email #
    time.sleep(1)
    pyautogui.write(email)
    pyautogui.press('enter')

    # Pular para Assunto #
    pyautogui.press('tab')
    pyautogui.write('Cobrança Mensalidade')

    # Pular para o Corpo do email #
    pyautogui.press('tab')
    texto = '''
    olá, Caloteiro
    
    Pague suas dividas, vagabundo. '''

    # Copiar e colar 'texto' #
    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')

    # Finalmente enviar Email #
    pyautogui.hotkey('ctrl', 'enter')

# Avisar o usuario que o email foi enviado #
pyautogui.alert('Email enviado com sucesso.')
