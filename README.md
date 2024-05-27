# automacaoWeb
 
## Automação Web de Pesquisa de Preços

Este projeto é uma automação para pesquisa de preços de produtos na web, utilizando Google Shopping e Buscapé. A automação busca por produtos específicos, verifica se estão dentro de uma faixa de preço desejada e exporta os resultados para um arquivo Excel. Além disso, envia um e-mail com os resultados encontrados.

# Requisitos

Python 3.x
Webdriver do Chrome (chromedriver)
Bibliotecas Python:
selenium
pandas
openpyxl
win32com.client (pywin32)

# Instalação

1 Instale as dependências:
pip install selenium pandas openpyxl pywin32

2 Baixe o webdriver do Chrome:
Faça o download do ChromeDriver compatível com a versão do seu navegador.
Extraia o arquivo e coloque o executável no seu PATH ou no mesmo diretório do projeto.

# Uso
1 Prepare a planilha de entrada:

Crie um arquivo Excel chamado buscas.xlsx com as colunas: Nome, Termos banidos, Preço mínimo, Preço máximo.
Exemplo de conteúdo:
Nome;	Termos banidos;	Preço mínimo;	Preço máximo
iphone 12 64 gb;	mini watch;	3000;	3500
rtx 3060;	zota galax;	4000;	4500

2 Execute o script:

- Utilize o Jupyter Notebook para executar o código ou transforme-o em um script Python.
- O código principal está dividido em várias células, com os seguintes passos:
    - Configuração do navegador.
    - Importação e visualização da base de dados.
    - Definição das funções de busca no Google Shopping e Buscapé.
    - Construção da lista de ofertas encontradas.
    - Exportação da base de ofertas para Excel.
    - Envio do e-mail com os resultados.

# Estrutura do Código
1. Configuração do Navegador

   from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd

# Criar o navegador
nav = webdriver.Chrome()

# Importar e visualizar a base de dados
tabela_produtos = pd.read_excel("buscas.xlsx")
display(tabela_produtos)

2. Funções de Busca
- Google Shopping

import time

def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    # Código da função busca_google_shopping
    # ...
    return lista_ofertas

- Buscapé

def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    # Código da função busca_buscape
    # ...
    return lista_ofertas

3. Construção da Lista de Ofertas

tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, "Nome"]
    termos_banidos = tabela_produtos.loc[linha, "Termos banidos"]
    preco_minimo = tabela_produtos.loc[linha, "Preço mínimo"]
    preco_maximo = tabela_produtos.loc[linha, "Preço máximo"]
    
    lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto', 'preco', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_google_shopping)
    
    lista_ofertas_buscape = busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_buscape)

display(tabela_ofertas)

4.  Exportação para Excel

tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel("Ofertas.xlsx", index=False)

5. Envio do E-mail

import win32com.client as win32

if len(tabela_ofertas.index) > 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'emaildodestinatario@gmail.com'
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f"""
    <p>Prezados,</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada. Segue tabela com detalhes</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att.,</p>
    """
    mail.Send()

nav.quit()


