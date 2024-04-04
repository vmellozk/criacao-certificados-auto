'''
Pegar dados de uma planilha excel e passá-los para uma imagem de um certificado.
Dados como: curso, nome, tipo da participação, data de início e final, carga horária e data de emissão.
'''

import openpyxl
from PIL import Image, ImageDraw, ImageFont

#Abrindo a planilha
planilha = openpyxl.load_workbook('planilha_alunos.xlsx')
pag_planilha = planilha['Sheet1']

#Cada célula contém a info de cada 
for indice, coluna in enumerate(pag_planilha.iter_rows(min_row=2)):
    curso = coluna[0].value
    nome = coluna[1].value
    participacao = coluna[2].value
    carga_horaria = coluna[5].value

    data_inicio = coluna[3].value
    data_final = coluna[4].value
    
    data_emissao = coluna[6].value

#Tipos de Fontes
    fonte_negrito = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

#Abrindo Imagem
    imagem = Image.open('./certificado_padrao.jpg')
    aplicando_texto = ImageDraw.Draw(imagem)

    aplicando_texto.text((1020, 827), nome, fill='black', font=fonte_negrito,)
    aplicando_texto.text((1060, 950), curso, fill='black', font=fonte,)
    aplicando_texto.text((1435, 1065), participacao, fill='black', font=fonte,)
    aplicando_texto.text((1480, 1182), str(carga_horaria), fill='black', font=fonte,)

    aplicando_texto.text((750, 1770), data_inicio, fill='black', font=fonte_data,)
    aplicando_texto.text((750, 1930), data_final, fill='black', font=fonte_data,)

    aplicando_texto.text((2220, 1930), data_emissao, fill='black', font=fonte_data,)


    imagem.save(f'./{indice + 1}_{nome}_[CERTIFICADO].jpg')