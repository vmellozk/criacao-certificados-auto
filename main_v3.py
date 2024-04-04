'''
Pegar dados de uma planilha excel e passá-los para uma imagem de um certificado.
Dados como: curso, nome, tipo da participação, data de início e final, carga horária e data de emissão.
'''

import openpyxl
from PIL import Image, ImageDraw, ImageFont

nomes_processados = []

#Abrindo a planilha
planilha = openpyxl.load_workbook('planilha_alunos.xlsx')
pag_planilha = planilha['Sheet1']

#Cada célula contém a info de cada 
for indice, coluna in enumerate(pag_planilha.iter_rows(min_row=2)):
    if None in coluna:
        print(f"Erro na linha{indice}: Valores em branco!")
        continue

    curso = coluna[0].value if coluna[0].value is not None else ""
    nome = coluna[1].value if coluna[1].value is not None else ""
    participacao = coluna[2].value if coluna[2].value is not None else ""
    carga_horaria = coluna[5].value if coluna[3].value is not None else ""

    data_inicio = coluna[3].value if coluna[4].value is not None else ""
    data_final = coluna[4].value if coluna[5].value is not None else ""
    
    data_emissao = coluna[6].value if coluna[6].value is not None else ""

    if nome in nomes_processados:
        print(f"Erro na linha {indice}: o nome já foi processado.")
        continue
    nomes_processados.append(nome)

#Tipos de Fontes
    fonte_negrito = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

#Abrindo Imagem
    imagem = Image.open('./certificado_padrao.jpg')
    aplicando_texto = ImageDraw.Draw(imagem)

#Transferindo os dados da planilha para a imagem do certificado
    aplicando_texto.text((1020, 827), nome, fill='black', font=fonte_negrito,)
    aplicando_texto.text((1060, 950), curso, fill='black', font=fonte,)
    aplicando_texto.text((1435, 1065), participacao, fill='black', font=fonte,)
    aplicando_texto.text((1480, 1182), str(carga_horaria), fill='black', font=fonte,)

    aplicando_texto.text((750, 1770), data_inicio, fill='black', font=fonte_data,)
    aplicando_texto.text((750, 1930), data_final, fill='black', font=fonte_data,)

    aplicando_texto.text((2220, 1930), data_emissao, fill='black', font=fonte_data,)


    imagem.save(f'./[{indice + 1}]{nome}_[CERTIFICADO].jpg')
