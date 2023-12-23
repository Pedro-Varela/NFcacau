import xml.etree.ElementTree as ET
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog
import functions as f


tk_window = tk.Tk()
tk_window.withdraw()  # para nao ter uma janela tkinter inteira aberta. Só quero a janela de seleção de arquivo 

file_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
print(file_path)  # Caminho para o arquivo XML selecionado




# Caminho do seu arquivo XML

# Carregar o arquivo XML
tree = ET.parse(file_path)
root = tree.getroot()

#root contém todo o xml

# Definir os namespaces
namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
arquivo_excel = 'dados_produtos.xlsx'

#ao usar o ElementTree, cada vez que se busca uma tag, o namespace tem que vir antes, tipo a url inteira para
#achar a tag especifica, então eu defino esse dicionario em python para poder chamalo sempre que eu quiser procurar uma tag


# Buscar um elemento específico, incluindo o namespace
numero_nota = root.find('.//nfe:nNF', namespaces).text

if f.verificar_numero_nota(numero_nota):
    print("Este arquivo XML já foi processado.")
else:
    # Processar o arquivo XML
    # ...
    f.registrar_numero_nota(numero_nota)

# Iterar sobre todos os elementos 'prod'
# Assumindo que você já definiu 'root' e 'namespaces' corretamente
produtos = []

for prod in root.findall('.//nfe:prod', namespaces):
    # Dicionário para armazenar os dados do produto atual
    dados_do_produto = {}

    # Extrair e armazenar cada dado do produto

    #Codigo do produto
    dados_do_produto['cProd'] = prod.find('nfe:cProd', namespaces).text.lstrip('0') if prod.find('nfe:cProd', namespaces) is not None else ''
    #Descrição do produto
    dados_do_produto['xProd'] = prod.find('nfe:xProd', namespaces).text if prod.find('nfe:xProd', namespaces) is not None else ''
    #Valor do produto
    dados_do_produto['vProd'] = prod.find('.//nfe:vProd', namespaces).text if prod.find('.//nfe:vProd', namespaces) is not None else ''
    #ICMS 
    
    #parei aqui, falta colocar ICMS e fazer o resto do codigo 
    # Adicionar os dados do produto à lista de produtos
    produtos.append(dados_do_produto)


    # Supondo que 'numero_nota' seja extraído do arquivo XML

    



f.adicionar_dados_ao_excel(arquivo_excel, numero_nota, produtos)




