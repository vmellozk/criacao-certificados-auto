# Certificados_Auto
 Este projeto consiste em um script em Python que automatiza a geração de certificados a partir de dados contidos em uma planilha Excel. O script lê as informações da planilha, como curso, nome do participante, tipo de participação, datas de início e término, carga horária e data de emissão, e as insere em um modelo de certificado pré-definido, gerando assim um certificado personalizado para cada participante do curso.

Funcionamento do Código:



Leitura de Dados da Planilha Excel:
O script utiliza a biblioteca openpyxl para ler os dados contidos na planilha Excel.
Itera sobre as linhas da planilha, coletando as informações relevantes de cada participante do curso.

Processamento e Validação de Dados:
Os dados são processados e validados para garantir que estejam completos e consistentes.
São verificados se há valores em branco e se o nome do participante já foi processado anteriormente.

Geração de Certificados Personalizados:
Utiliza a biblioteca PIL (Python Imaging Library) para abrir o modelo de certificado pré-definido em formato de imagem.
As informações coletadas da planilha são então inseridas no certificado, utilizando diferentes fontes e coordenadas para cada tipo de informação.

Salvamento dos Certificados Gerados:
Após a inserção das informações, o certificado personalizado é salvo como uma nova imagem JPG, com o nome do participante e um identificador numérico para diferenciar cada certificado.

Tratamento de Exceções:
O script trata exceções como arquivos não encontrados e erros inesperados, garantindo que o processo de geração de certificados seja robusto e confiável.



Explicação dos Métodos Utilizados:

openpyxl.load_workbook(): Utilizado para carregar a planilha Excel e acessar suas células.
Image.open(): Utilizado para abrir o modelo de certificado em formato de imagem.
ImageDraw.Draw(): Utilizado para desenhar texto sobre a imagem do certificado.
ImageFont.truetype(): Utilizado para carregar fontes TrueType para inserir texto no certificado.
imagem.save(): Utilizado para salvar os certificados gerados como novas imagens JPG.
