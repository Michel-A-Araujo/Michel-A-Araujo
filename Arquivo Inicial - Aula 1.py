#!/usr/bin/env python
# coding: utf-8

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

# In[85]:


import pyautogui
import pyperclip
import time

# pyautogui.click
# pyautogui.write
# pyautogui.press
# pyautogui.hotkey


# Passo 1: Entrar no site da empresa (https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing)
pyautogui.hotkey ("ctrl","t")
pyperclip.copy ("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey ("ctrl","v")
pyautogui.press ("enter")   
time.sleep (5)

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=1016, y=293, clicks=2)
time.sleep (2)

# Passo 3: Exportar o relatório (fazer o download)
pyautogui.rightClick (x=1042, y=285)
time.sleep (2)
pyautogui.click(x=1081, y=631)


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[86]:


# Passo 4: Calcular os Indicadores (Faturamento e quantidade de produtos vendidos)
import pandas

tabela = pandas.read_excel(r"C:\Users\ianno\Downloads\Vendas - Dez.xlsx")
display(tabela)

quantidade = tabela["Quantidade"].sum()
faturamento = tabela["Valor Final"].sum()
print(quantidade)
print(faturamento)


# ### Vamos agora enviar um e-mail pelo gmail

# In[87]:


# Passo 5: Enviar o e-mail para a diretoria (seugmail+diretoria@gmail.com)
# abrir aba
pyautogui.hotkey ("ctrl","t")

# entrar n e-mail
pyperclip.copy ("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey ("ctrl","v")
pyautogui.press ("enter")   
time.sleep (5)

# clicar no botão escrever
pyautogui.click (x=755, y=212)

# preencher as informações do e-mail
pyperclip.copy ("iannottingham@msn.com")
pyautogui.hotkey ("ctrl","v")  
pyautogui.press ("tab")   
pyperclip.copy ("Relatório de Vendas")
pyautogui.hotkey ("ctrl","v")
pyautogui.press ("tab")   

texto = f"""
Prezados senhores,

O faturamento deste mês fechou em R$ {faturamento:,.2f} e no total foram vendidos {quantidade:,} itens.

Atenciosamente,

Michel Araujo"""
pyperclip.copy(texto)
pyautogui.hotkey ("ctrl","v")           
                  
# enviar o email
pyautogui.hotkey ("ctrl","enter")


# Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[88]:


import time
time.sleep (5)
pyautogui.position()


# In[ ]:





# In[ ]:





# In[ ]:




