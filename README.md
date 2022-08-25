  #Importar bibliotecas
import pyautogui
import pyperclip
import time
import pandas as pd


  #Entrar no sistema(neste caso, entrar no link)
  
pyautogui.hotkey('ctrl', 't')


  #Navegar até o local do relatório 
  
pyautogui.write('https://drive.google.com/drive/folders/14oLE59U1RqyRqlBbKpsyymW-mitvbtoh')
pyautogui.press('enter')
time.sleep(2)

  #Fazer o dowload do relatório
  
pyautogui.click(x=389, y=342)
time.sleep(1)
pyautogui.click(x=1161, y=159)
time.sleep(1)
pyautogui.click(x=957, y=567)

  #ler o arquivo baixado para pegar os indicadores
  
tabela = pd.read_excel(r'Vendas - Dez.xlsx')

#Visualizar a tabela

display(tabela)

  #Calcular os indicadores
  
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

  #visualizar os indicadores
  
display(faturamento)
display(quantidade)

    #enviar um e-mail pelo gmail
    
  #Entrar no email
  
pyautogui.hotkey('crtl', 't')
time.sleep(2)
pyautogui.click(x=1160, y=100)
time.sleep(2)
  #Enviar por email o resultado
pyautogui.click(x=75, y=174)
time.sleep(1)

pyautogui.write('Treinandopython06112002@gmail.com')
pyautogui.press('tab') #seleciona o email
pyautogui.press('tab') #pula para o campo de assunto
time.sleep(1)
pyperclip.copy('Relatório de vendas')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab') #pular para o corpo do email
time.sleep(1)

texto = f'''Prezados, bom dia


O faturamento de ontem foi de: {faturamento:,.2f}
A quantidade de produtos foi de: {quantidade}

Abs
Gabriel'''

time.sleep(1)
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)

  #clicar no botão enviar
  
pyautogui.hotkey('ctrl', 'enter')

