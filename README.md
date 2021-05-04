import pyautogui
import time
import pyperclip


# In[36]:


pyautogui.PAUSE = 2
pyautogui.alert(" O SHOW VAI COMEÇAR, DEIXA COMIGO")
pyautogui.hotkey('ctrl','t')
link = "https://drive.google.com"
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(8)
pyautogui.click(626,708)
pyautogui.click(1165,161)
pyautogui.click(1022,566)
time.sleep(10)


# In[37]:


import pandas as pd
df = pd.read_excel(r"C:\Users\........")
display(df)
variaveis=  df ['Valores1'] .sum()
fixo = df ['Valores2'] .sum()


# In[38]:


pyautogui.hotkey('ctrl','t')
link = "mail.google.com"
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(8)
pyautogui.click(72,175)
time.sleep(4)
pyautogui.write('@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = "Relatório de Despesas"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')
text_email = f""" Olá Prezados

O valor das Despesas Variáveis desta semana foi de: R$ {variaveis:,.2f}

O valor das Despesas Fixas desta semana foi de: R$ {fixo:,.2f}


Obrigado"""
pyperclip.copy(text_email)
pyautogui.hotkey('ctrl','v')
time.sleep(3)
pyautogui.hotkey('ctrl','enter')


# In[ ]:





