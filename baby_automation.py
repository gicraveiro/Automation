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
Oi!

Aqui quem fala é o computador da Giovana, acabei de ser promovido a secretário com independência para enviar meus próprios e-mails! Só não posso mesmo é escolher o que está escrito.

No primeiro exercício que ela fez para aprender a automatizar alguns processos, ela criou um código para que eu conseguisse sozinho fazer tudo isso:
1 - Abrir o navegador
2 - Fazer o download de uma base de dados que estava salva em uma pasta do seu drive
3 - Somar a quantidade de produtos e o faturamento total da base de dados (que são esses valores no fim do e-mail)
4 - Enviar este e-mail

Quantidade de produtos: {quantidade:,}
Faturamento: {faturamento:,.2f}

Ela queria que eu passasse algumas observações sobre o código também:
-> Eu estou um pouco lentinho, por isso que o código ta cheio de linhas pedindo pro próximo comando esperar uns 2 segundos antes de rodar.
-> As posições dos cliques, o acesso aos arquivos sem precisar da senha, e outras características específicas dessa fazem com que esse código funcione só quando eu é que estou rodando. Mas fica tranquilo que é só alterar algumas coisas quando algum outro pc quiser rodar.

Atenciosamente,

Computador da Giovana
 """
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
time.sleep(2)
pyautogui.click(x=321,y=625)
