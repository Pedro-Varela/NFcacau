import xml.etree.ElementTree as ET
import openpyxl
import os


# Caminho do seu arquivo XML
xml_file = '001206201.xml'

# Carregar o arquivo XML
tree = ET.parse(xml_file)
root = tree.getroot()

#root contém todo o xml

# Definir os namespaces
namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
arquivo_excel = 'dados_produtos.xlsx'

#ao usar o ElementTree, cada vez que se busca uma tag, o namespace tem que vir antes, tipo a url inteira para
#achar a tag especifica, então eu defino esse dicionario em python para poder chamalo sempre que eu quiser procurar uma tag


# Buscar um elemento específico, incluindo o namespace
numero_nota = root.find('.//nfe:nNF', namespaces).text

# Iterar sobre todos os elementos 'prod'
# Assumindo que você já definiu 'root' e 'namespaces' corretamente
produtos = []

for prod in root.findall('.//nfe:prod', namespaces):
    # Dicionário para armazenar os dados do produto atual
    dados_do_produto = {}

    # Extrair e armazenar cada dado do produto
    dados_do_produto['cProd'] = prod.find('nfe:cProd', namespaces).text.lstrip('0d') if prod.find('nfe:cProd', namespaces) is not None else ''
    dados_do_produto['xProd'] = prod.find('nfe:xProd', namespaces).text if prod.find('nfe:xProd', namespaces) is not None else ''
    dados_do_produto['vProd'] = prod.find('nfe:vProd', namespaces).text if prod.find('nfe:vProd', namespaces) is not None else ''
    # Repita para outros campos conforme necessário

    # Adicionar os dados do produto à lista de produtos
    produtos.append(dados_do_produto)



def adicionar_dados_ao_excel(arquivo_excel, numero_nota, dados_produtos):
    # Verificar se o arquivo Excel existe
    if os.path.exists(arquivo_excel):
        workbook = openpyxl.load_workbook(arquivo_excel)
    else:
        workbook = openpyxl.Workbook()
        # Adicionar cabeçalhos se for um novo arquivo
        headers = ['Número da Nota', 'Código do Produto', 'Descrição do Produto', 'Valor do Produto']
        workbook.active.append(headers)

    sheet = workbook.active

    # Adicionar os dados dos produtos
    for produto in dados_produtos:
        row = [
            numero_nota, 
            produto['cProd'],  # Certificando que é uma string
            produto['xProd'],  # Certificando que é uma string
            produto['vProd']   # Certificando que é uma string
        ]
        sheet.append(row)

    # Salvar o arquivo Excel
    workbook.save(arquivo_excel)


adicionar_dados_ao_excel(arquivo_excel, numero_nota, produtos)




