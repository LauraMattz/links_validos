import openpyxl
import requests

# abrir o arquivo Excel
workbook = openpyxl.load_workbook('Pasta1.xlsx')

# selecionar a planilha
sheet = workbook.active

# criar uma lista para armazenar os links inválidos
links_invalidos = []

# iterar sobre as linhas da planilha
for row in sheet.iter_rows(min_row=2, min_col=2, max_col=2):
    # obter o valor da célula
    valor = row[0].value

    # verificar se a célula contém um valor
    if valor:
        # verificar se o valor é uma string e começa com uma URL
        if isinstance(valor, str) and (valor.startswith('http') or valor.startswith('https')):
            # enviar uma solicitação HTTP ao link
            resposta = requests.get(valor)

            # verificar se a resposta tem um código de status válido
            if resposta.status_code != 200:
                # adicionar o link inválido à lista
                links_invalidos.append(valor)

# imprimir a lista de links inválidos
if links_invalidos:
    print('Os seguintes links são inválidos:')
    for link in links_invalidos:
        print(link)
else:
    print('Todos os links são válidos.')
