import os
import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

# Open browser
pyautogui.press("win")
time.sleep(2)
pyautogui.click(x=567,y=69)
pyautogui.write("brave")
pyautogui.press("enter") 

# Open drive, download chart
time.sleep(2)
pyautogui.hotkey("ctrl", "t") # open a new tab
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter") 
time.sleep(2)
pyautogui.click(x=347,y=300,clicks=2) 
time.sleep(2)
pyautogui.click(x=347,y=300) 
time.sleep(2)
pyautogui.click(x=1161,y=186)
time.sleep(2)
pyautogui.click(x=904,y=557)
time.sleep(2)
pyautogui.press("enter") 
time.sleep(5)

# Read chart
tabela = pd.read_excel(r"/home/gi/Downloads/Vendas - Dez.xlsx")
#display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento, quantidade)

pyautogui.hotkey("ctrl", "t") # open a new tab
pyperclip.copy("https://outlook.live.com/mail/0/")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=183,y=165)
time.sleep(2)
pyautogui.write("eng.flavia.firmino@gmail.com")
pyautogui.press("enter")
pyautogui.write("gicraveiroo@gmail.com")
pyautogui.press("enter")
pyautogui.press("tab")
pyautogui.write("Python responde")
pyautogui.press("tab")
texto = f"""
Oi Fla!

Aqui quem fala é o computador da LG, acabei de ser promovido a secretário com independência para enviar meus próprios e-mails! Só não posso mesmo é escolher o que está escrito.
A LG adorou a mensagem e queria te informar que ela finalmente conseguiu completar o exercício!

Eu tbm to com saudades de ter um descanso enquanto ela joga jogos, vai no mercado e cuida das plantas com vc, triste :/
Não tem nenhuma louça dela na pia, não, viu (até onde ela sabe rs)

Só pra checar que ta tudo certo mesmo, esses são os números do relatório:
Quantidade de produtos: {quantidade:,}
Soma total: {faturamento:,.2f}

Atenciosamente,

Computador da LG
 """
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
time.sleep(2)
pyautogui.click(x=321,y=625)
