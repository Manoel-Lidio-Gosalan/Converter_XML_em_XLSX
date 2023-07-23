import xmltodict
import os
import pandas as pd
import json


def pegar_infos(nome_arquivo, valores):
    print(f'Peguei as informações {nome_arquivo}')
    with open(f'xmls/{nome_arquivo}', 'rb') as arquivo_xml:
        dict_arquivo = xmltodict.parse(arquivo_xml)
        
        if 'NFe' in dict_arquivo:
            info_nf = dict_arquivo['NFe']['infNFe']
        else:
            info_nf = dict_arquivo['nfeProc']['NFe']['infNFe']
        
        numero_nf = info_nf['@Id']
        empresa_emissora = info_nf['emit']['xNome']
        nome_cliente = info_nf['dest']['xNome']
        endereco_cliente = info_nf['dest']['enderDest']['xLgr']
        numero_endereco_cliente = info_nf['dest']['enderDest']['nro']
        
        if 'xCpl' in info_nf['dest']:
            complemento_endereco_cliente = info_nf['dest']['enderDest']['xCpl']
        else:
            complemento_endereco_cliente = 'Não informado o complemento do endereço'
        
        bairro_endereco_cliente = info_nf['dest']['enderDest']['xBairro']
        cep_endereco_cliente = info_nf['dest']['enderDest']['CEP']
        cidade_cliente = info_nf['dest']['enderDest']['xMun']
        uf_cliente = info_nf['dest']['enderDest']['UF']
        
        if 'vol' in info_nf['transp']:
            Peso_Bruto = info_nf['transp']['vol']['pesoB']
        else:
            Peso_Bruto = 'Não informado o peso bruto na nota'
        
        valores.append([numero_nf, empresa_emissora, Peso_Bruto, nome_cliente, endereco_cliente,numero_endereco_cliente,complemento_endereco_cliente, bairro_endereco_cliente, cep_endereco_cliente, cidade_cliente, uf_cliente])


lista_arquivos = os.listdir('xmls')

colunas = ['numero_nf', 'empresa_emissora', 'Peso_Bruto','nome_cliente', 'endereço_cliente', 'numero_endereço_cliente', 'complemento_endereço_cliente', 'bairro_endereço_cliente', 'cep_endereço_cliente', 'cidade_cliente', 'uf_cliente']
valores = []

for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)
    

tabela = pd.DataFrame(columns=colunas, data=valores)

tabela.to_excel('notas_fiscais.xlsx', index=False)
