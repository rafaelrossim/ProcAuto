# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %% [markdown]
# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado
# 
# Comandos pyautogui: https://pyautogui.readthedocs.io/en/latest/quickstart.html

# %%
# 1 - escrever passo a passo o que você faria manualmente


# %%
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.sleep(5)
pyautogui.click(x=571, y=409)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=438, y=655, clicks=2)
time.sleep(2)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=421, y=583, button='right')
pyautogui.click(x=490, y=643)
time.sleep(5)
pyautogui.click(x=1048, y=123)
pyautogui.write("D:\Projetos\Python\PythonIntensivao\docs\Aula1")
pyautogui.press("enter")
pyautogui.click(x=1190, y=687)
pyautogui.click(x=1347, y=695)


# %%
# Passo 4: Calcular os indicadores
import pandas as pd

tabela = pd.read_excel(r"D://Projetos/Python/PythonIntensivao/docs/Aula1/Vendas - Dez.xlsx")
# se a planilha tiver mais de uma aba
#tabela = pd.read_excel(r"D://Projetos/Python/PythonIntensivao/docs/Vendas - Dez.xlsx", sheets="nome da aba ou numero")
#display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

# %% [markdown]
# ### Vamos agora enviar um e-mail pelo gmail

# %%
# Passo 5: Entrar no email
time.sleep(5)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=101, y=259)

pyautogui.write("beatriz-kellyribeiro@outlook.com")
pyautogui.press("tab") # assunto
pyautogui.press("tab") 
pyperclip.copy("Resumo de vendas - Mensal") # escrever o assunto
pyautogui.hotkey("ctrl", "v")   
pyautogui.press("tab") #pular pro corpo do email

texto = f"""
Boa noite

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
RafaelRossim"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.click(x=754, y=691)
pyautogui.click(x=1044, y=122)
pyautogui.write("D:\Projetos\Python\PythonIntensivao\docs\Aula1")
pyautogui.press("enter")
pyautogui.click(x=279, y=259, clicks=2)
pyautogui.sleep(5)

# apertar enter para enviar o email
pyautogui.hotkey("ctrl", "enter") # enviar email

# %% [markdown]
# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# %%
#time.sleep(5)
#pyautogui.position()


# %%
# Como instalar
#!pip install pyautogui
#!pip install pyperclip


