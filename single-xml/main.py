import xml.etree.ElementTree as ET
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
arquivo_excel_produtos  = 'dados_produtos.xlsx'
arquivo_excel_nota = 'dados_notas.xlsx'

#ao usar o ElementTree, cada vez que se busca uma tag, o namespace tem que vir antes, tipo a url inteira para
#achar a tag especifica, então eu defino esse dicionario em python para poder chamalo sempre que eu quiser procurar uma tag

# Buscar um elemento específico, incluindo o namespace
numero_nota = root.find('.//nfe:nNF', namespaces).text

if f.verificar_numero_nota(numero_nota):
    print("Este arquivo XML já foi processado.")
    exit()
else:
    # Processar o arquivo XML
    f.registrar_numero_nota(numero_nota)

# Iterar sobre todos os elementos 'prod' e extrair os dados
produtos = []

#loop para pegar os dados dos produtos, que esta dentro da tag prod
for prod in root.findall('.//nfe:prod', namespaces):
    # Dicionário para armazenar os dados do produto atual
    dados_do_produto = {}
    # Extrair e armazenar cada dado do produto
    #Ean do produto
    dados_do_produto['cEAN'] = prod.find('nfe:cEAN', namespaces).text if prod.find('nfe:cEAN', namespaces) is not None else ''
    #NCM do produto
    dados_do_produto['NCM'] = prod.find('nfe:NCM', namespaces).text if prod.find('nfe:NCM', namespaces) is not None else ''
    #Cest do produto
    dados_do_produto['CEST'] = prod.find('nfe:CEST', namespaces).text if prod.find('nfe:CEST', namespaces) is not None else ''
    #Codigo do produto
    dados_do_produto['cProd'] = prod.find('nfe:cProd', namespaces).text.lstrip('0') if prod.find('nfe:cProd', namespaces) is not None else ''
    #Descrição do produto
    dados_do_produto['xProd'] = prod.find('nfe:xProd', namespaces).text if prod.find('nfe:xProd', namespaces) is not None else ''
    #Quantidade de Caixas  
    dados_do_produto['qCom'] = prod.find('.//nfe:qCom', namespaces).text.rstrip('.0') if prod.find('.//nfe:qCom', namespaces) is not None else ''
    #Valor unitario do produto
    dados_do_produto['vUnCom'] = prod.find('.//nfe:vUnCom', namespaces).text if prod.find('.//nfe:vUnCom', namespaces) is not None else ''
    #Valor total do produto
    dados_do_produto['vProd'] = prod.find('.//nfe:vProd', namespaces).text if prod.find('.//nfe:vProd', namespaces) is not None else ''
    #Numero do lote 
    dados_do_produto['nLote'] = prod.find('.//nfe:nLote', namespaces).text if prod.find('.//nfe:nLote', namespaces) is not None else ''
    #Data de Fabricacao
    dados_do_produto['dFab'] = prod.find('.//nfe:dFab', namespaces).text if prod.find('.//nfe:dFab', namespaces) is not None else ''
    #Data de Validade
    dados_do_produto['dVal'] = prod.find('.//nfe:dVal', namespaces).text if prod.find('.//nfe:dVal', namespaces) is not None else ''
   # Adicionar os dados do produto à lista de produtos
    produtos.append(dados_do_produto)

# Buscar os valores de vICMS
vICMS = []
pICMS = []
vIPI = []
pIPI = []
vPIS = []
pPIS = []
vCOFINS = []
pCOFINS = []

#loop para pegar os dados dos impostos, que esta dentro da tag imposto.
for imposto in root.findall('.//nfe:imposto', namespaces):
    vICM = imposto.find('.//nfe:vICMS', namespaces).text
    vICMS.append(vICM)

    pICM = imposto.find('.//nfe:pICMS', namespaces).text
    pICMS.append(pICM)

    # tem essa verificação para caso o produto não tenha IPI, ele não quebre o programa, considerando que poucos produtos tem ipi
    vIPI_tag = imposto.find('.//nfe:vIPI', namespaces)
    if vIPI_tag is not None:
        vIPI_value = vIPI_tag.text
    else:
        vIPI_value = '0'  
    vIPI.append(vIPI_value)  # Adiciona o valor à lista vIPI

    # Para alíquota do IPI (pIPI)
    pIPI_tag = imposto.find('.//nfe:pIPI', namespaces)
    if pIPI_tag is not None:
        pIPI_value = pIPI_tag.text
    else:
        pIPI_value = '0'  # ou qualquer valor padrão que represente "sem alíquota de IPI"
    pIPI.append(pIPI_value)  # Adiciona o valor à lista p

    vPI = imposto.find('.//nfe:vPIS', namespaces).text
    vPIS.append(vPI)

    pPI = imposto.find('.//nfe:pPIS', namespaces).text
    pPIS.append(pPI)

    vCO = imposto.find('.//nfe:vCOFINS', namespaces).text
    vCOFINS.append(vCO)

    pCO = imposto.find('.//nfe:pCOFINS', namespaces).text
    pCOFINS.append(pCO)

# Adicionar os valores de vICMS aos produtos
# Mapeamento das chaves para as respectivas listas
# antes era varios IFS, criei esse dicionario para ficar mais facil de adicionar os valores e mais inteligente doq varios ifs.

impostos_map = {
    'vICMS': vICMS,
    'pICMS': pICMS,
    'vIPI': vIPI,
    'pIPI': pIPI,
    'vPIS': vPIS,
    'pPIS': pPIS,
    'vCOFINS': vCOFINS,
    'pCOFINS': pCOFINS
}

for i, produto in enumerate(produtos):
    for chave, lista in impostos_map.items():
        produto[chave] = lista[i] if i < len(lista) else ''





# Criar o dicionário fora dos loops para acumular dados da nota
dados_da_nota = {}
infonota = []
# Loop para extrair informações da tag 'ide'
for info in root.findall('.//nfe:ide', namespaces):
    # Preenchendo dados da nota
    dados_da_nota['dhEmi'] = info.find('nfe:dhEmi', namespaces).text if info.find('nfe:dhEmi', namespaces) is not None else ''
    dados_da_nota['dhSaiEnt'] = info.find('nfe:dhSaiEnt', namespaces).text if info.find('nfe:dhSaiEnt', namespaces) is not None else ''
    dados_da_nota['cMunFG'] = info.find('nfe:cMunFG', namespaces).text if info.find('nfe:cMunFG', namespaces) is not None else ''
    dados_da_nota['tpImp'] = info.find('nfe:tpImp', namespaces).text if info.find('nfe:tpImp', namespaces) is not None else ''

# Loop para extrair informações da tag 'emit'
for info in root.findall('.//nfe:emit', namespaces):
    # Preenchendo dados da nota
    dados_da_nota['CNPJ_emit'] = info.find('nfe:CNPJ', namespaces).text if info.find('nfe:CNPJ', namespaces) is not None else ''  # Mudança no nome da chave para evitar sobreposição
    dados_da_nota['xNome_emit'] = info.find('nfe:xNome', namespaces).text if info.find('nfe:xNome', namespaces) is not None else ''
    dados_da_nota['xFant_emit'] = info.find('nfe:xFant', namespaces).text if info.find('nfe:xFant', namespaces) is not None else ''
    dados_da_nota['IE_emit'] = info.find('nfe:IE', namespaces).text if info.find('nfe:IE', namespaces) is not None else ''
    dados_da_nota['CNAE_emit'] = info.find('nfe:CNAE', namespaces).text if info.find('nfe:CNAE', namespaces) is not None else ''


for info in root.findall('.//nfe:enderEmit', namespaces):
    dados_da_nota['xMun_emit'] = info.find('nfe:xMun', namespaces).text if info.find('nfe:xMun', namespaces) is not None else ''
    dados_da_nota['UF_emit'] = info.find('nfe:UF', namespaces).text if info.find('nfe:UF', namespaces) is not None else ''
    
# Loop para extrair informações da tag 'totais'
for info in root.findall('.//nfe:ICMSTot', namespaces):
    # Preenchendo dados da nota
    dados_da_nota['vNF'] = info.find('nfe:vNF', namespaces).text if info.find('nfe:vNF', namespaces) is not None else ''
    dados_da_nota['vProd_totais'] = info.find('nfe:vProd', namespaces).text if info.find('nfe:vProd', namespaces) is not None else ''  # Mudança no nome da chave para evitar sobreposição
    dados_da_nota['vICMS_totais'] = info.find('nfe:vICMS', namespaces).text if info.find('nfe:vICMS', namespaces) is not None else ''
    dados_da_nota['vPIS_totais'] = info.find('nfe:vPIS', namespaces).text if info.find('nfe:vPIS', namespaces) is not None else ''
    dados_da_nota['vCOFINS_totais'] = info.find('nfe:vCOFINS', namespaces).text if info.find('nfe:vCOFINS', namespaces) is not None else ''

# Loop para extrair informações da tag 'transporta'
for info in root.findall('.//nfe:transporta', namespaces):
    # Preenchendo dados da nota
    dados_da_nota['xNome_transp'] = info.find('nfe:xNome', namespaces).text if info.find('nfe:xNome', namespaces) is not None else ''  # Mudança no nome da chave para evitar sobreposição
    dados_da_nota['CNPJ_transp'] = info.find('nfe:CNPJ', namespaces).text if info.find('nfe:CNPJ', namespaces) is not None else ''  # Mudança no nome da chave para evitar sobreposição

# Agora 'dados_da_nota' contém todas as informações reunidas
    
infonota.append(dados_da_nota)

print(dados_da_nota)
# Adicionar os dados do produto ao arquivo Excel
f.adicionar_dados_ao_excel_produtos(arquivo_excel_produtos, numero_nota, produtos)


f.adicionar_dados_ao_excel_notas(arquivo_excel_nota, numero_nota, infonota)


