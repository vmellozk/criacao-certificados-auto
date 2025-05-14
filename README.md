# 🏅 Certificados_Auto

## 📋 Descrição

Este projeto em Python automatiza a geração de certificados personalizados a partir de dados contidos em uma planilha Excel. O script insere as informações de cada participante (nome, curso, tipo de participação, datas, carga horária etc.) em um modelo de certificado pré-definido e salva o resultado como uma imagem.

## ⚙️ Tecnologias Utilizadas

- 📊 `openpyxl`: Leitura da planilha Excel com os dados dos participantes.
- 🖼️ `PIL (Pillow)`: Manipulação do modelo de certificado e inserção de texto.
- 💾 `os`: Acesso e manipulação de arquivos do sistema.
- 🛠️ Tratamento de exceções: Para garantir robustez ao processo.

## 🚀 Como Usar

### 1. Instalação das Dependências

Instale as bibliotecas necessárias com:

```bash
pip install openpyxl pillow
```

### 2. Preparação dos Arquivos

- Planilha Excel com os dados dos participantes.
- Modelo de certificado em formato .jpg.
- Fontes TrueType (.ttf) para estilização dos textos no certificado.

Certifique-se de que todos esses arquivos estejam no mesmo diretório do script.

### 3. Executando o Script

Rode o script com o seguinte comando:

```bash
python certificados_auto.py
```

Os certificados gerados serão salvos como imagens .jpg, com nomes personalizados baseados no nome do participante e um identificador numérico.

## 🧠 Funcionamento

### 1. Leitura da Planilha: O script lê os dados da planilha usando openpyxl.
### 2. Validação dos Dados: Verifica se há campos vazios ou duplicados.
### 3. Geração dos Certificados:
### 4. Abre o modelo de certificado com PIL.
### 5. Insere os dados nos locais corretos usando fontes específicas.
### 6. Salvamento das Imagens: Cada certificado é salvo como uma nova imagem .jpg.
### 7. Tratamento de Erros: Exceções como arquivos ausentes ou erros inesperados são tratados para evitar falhas na execução.

## 🔧 Principais Métodos Utilizados

- openpyxl.load_workbook() → Carrega a planilha Excel.
- Image.open() → Abre o modelo de certificado.
- ImageDraw.Draw() → Permite desenhar (escrever) sobre a imagem.
- ImageFont.truetype() → Carrega uma fonte personalizada.
- imagem.save() → Salva o certificado final como imagem.

## 📄 Licença
Este projeto está licenciado sob a MIT License. O arquivo 'LICENSE' ainda será disponibilizado para mais informações.
