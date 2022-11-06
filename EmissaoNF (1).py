#!/usr/bin/env python
# coding: utf-8

# In[10]:


from selenium import webdriver

from selenium.webdriver.common.by import By

options = webdriver.ChromeOptions()
navegador = webdriver.Chrome(options=options)


# In[11]:


# entrar na página de login (no nosso caso é login.html)
import os

caminho = os.getcwd()
arquivo = caminho + r"\index.html"
navegador.get(arquivo)


# In[12]:


# importar a base de clientes
import pandas as pd

tabela = pd.read_excel("NotasEmitir.xlsx") 
display(tabela)


# In[13]:


# para cada cliente - rodar o processo de emissao de nota fiscal
for linha in tabela.index:
    # preencher os dados da NF
    
    # nome/razao social
    navegador.find_element(By.NAME, 'nome').send_keys(tabela.loc[linha, "Cliente"])

    # endereco
    navegador.find_element(By.NAME, 'endereco').send_keys(tabela.loc[linha, "Endereço"])

    # bairro
    navegador.find_element(By.NAME, 'bairro').send_keys(tabela.loc[linha, "Bairro"])

    # municipio
    navegador.find_element(By.NAME, 'municipio').send_keys(tabela.loc[linha, "Municipio"])

    # cep
    navegador.find_element(By.NAME, 'cep').send_keys(str(tabela.loc[linha, "CEP"]))
    
    # UF
    navegador.find_element(By.NAME, 'uf').send_keys(tabela.loc[linha, "UF"])
    
    # CPF/CNPJ
    navegador.find_element(By.NAME, 'cnpj').send_keys(str(tabela.loc[linha, "CPF/CNPJ"]))

    # Inscricao estadual
    navegador.find_element(By.NAME, 'inscricao').send_keys(str(tabela.loc[linha, "Inscricao Estadual"]))

    # descrição
    texto = tabela.loc[linha, "Descrição"]
    navegador.find_element(By.NAME, 'descricao').send_keys(texto)

    # quantidade
    navegador.find_element(By.NAME, 'quantidade').send_keys(str(tabela.loc[linha, "Quantidade"]))

    # valor unitario
    navegador.find_element(By.NAME, 'valor_unitario').send_keys(str(tabela.loc[linha, "Valor Unitario"]))

    # valor total
    navegador.find_element(By.NAME, 'total').send_keys(str(tabela.loc[linha, "Valor Total"]))
    
    # clicar em emitir nota fiscal
    navegador.find_element(By.CLASS_NAME, 'registerbtn').click()
    
    # recarregar a página para limpar o formulário
    navegador.refresh()


# In[14]:


navegador.quit()


# In[ ]:




