import xml.etree.ElementTree as ET
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog



def registrar_numero_nota(numero_nota, arquivo_registro='notas_processadas.txt'):
    with open(arquivo_registro, 'a') as file:
        file.write(numero_nota + '\n')

def verificar_numero_nota(numero_nota, arquivo_registro='notas_processadas.txt'):
    try:
        with open(arquivo_registro, 'r') as file:
            notas_processadas = file.read().splitlines()
            return numero_nota in notas_processadas
    except FileNotFoundError:
        return False

def adicionar_dados_ao_excel(arquivo_excel, numero_nota, dados_produtos):
    # Verificar se o arquivo Excel existe
    if os.path.exists(arquivo_excel):
        workbook = openpyxl.load_workbook(arquivo_excel)
    else:
        workbook = openpyxl.Workbook()
        # Adicionar cabeçalhos se for um novo arquivo
        headers = ['Número da Nota', 'Código do Produto', 'Descrição do Produto', 'Valor do Produto', 'ICMS']
        workbook.active.append(headers)

    sheet = workbook.active

    # Adicionar os dados dos produtos
    for produto in dados_produtos:
        row = [
            numero_nota, 
            produto['cProd'],  # Certificando que é uma string
            produto['xProd'],  # Certificando que é uma string
            produto['vProd'],  # Certificando que é uma string
           # produto['vBC']    # Certificando que é uma string
        ]
        sheet.append(row)

    # Salvar o arquivo Excel
    workbook.save(arquivo_excel)
