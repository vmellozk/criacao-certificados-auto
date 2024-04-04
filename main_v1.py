'''
Pegar dados de uma planilha excel e passá-los para uma imagem de um certificado.
Dados como: curso, nome, tipo da participação, data de início e final, carga horária e data de emissão.
'''

import openpyxl
from PIL import Image, ImageDraw, ImageFont

#Abrindo a planilha
planilha = openpyxl.load_workbook('planilha_alunos.xlsx')
pag_planilha = planilha['pag1']

#Cada célula contém a info de cada 
for linha in pag_planilha.iter_rows(min_row=2):
    curso = linha[0].value
    nome = linha[1].value
    participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

#Tipos de Fontes
fonte_negrito = ImageFont.truetype('./tahomabd.ttf')
fonte = ImageFont.truetype('./tahoma.ttf')

#Abrindo Imagem
imagem = Image.open('./certificado_padrao.jpg')
aplicando_texto = ImageDraw.Draw(imagem)

aplicando_texto.text((1030, 960), nome, fill='black', font=fonte_negrito)
imagem.save('./teste.jpg')