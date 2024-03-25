################################################################
### Buscar cotação do dolar
################################################################

# Importar a biblioteca para pesquisarmos o cotação do dólar para o real, esta 
# biblioteca se chamada “yf”.
import yfinance as yf

# Usando a biblioteca declarada “yf” para recuperar a cotação USD/BRL
# Esta requisição recebe uma lista atual de cotações do dólar para real, 
# durante o dia atual
usd_brl = yf.Ticker('USDBRL=X')

# Obter o preço mais atual do par de moedas
usd_rate = usd_brl.history(period='1d')['Close'][0]

# Imprimir o valor da cotação do dolar
print(f'Cotação do dólar: USD 1 = R${usd_rate}')


################################################################
### Buscar criptomoeadas
################################################################

# Biblioteca que irá chamar o endereço onde buscaremos as cotações das criptomoeadas.
import requests

################################
### Código para buscar a cotação do Bitcoin para Real


# Moeda que iremos cotar
moedaCotacaoDe = 'bitcoin'

# A cotação será de BTC/BRL
moedaCotacaoPara = 'brl'

# Link html que irá retornar a cotação que queremos.
# Você pode colar esse link no seu navegador e você verá que os valores serão retornados na sua tela :)
# link: https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=brl
url = 'https://api.coingecko.com/api/v3/simple/price?ids='+ moedaCotacaoDe +'&vs_currencies=' + moedaCotacaoPara

# Realizar a chamada do link acima.
response = requests.get(url)

# Receber a resposta no format json (Você só precisa saber que json é um padrão de resposta)
data = response.json()

# Retirar o dado da resposta da variável "data" e coloca-la dentro da variavel "bitcoin_price_brl"
bitcoin_price_brl = data[moedaCotacaoDe][moedaCotacaoPara]

print(f'Cotação do Bitcoin é: R$ {bitcoin_price_brl:.2f}')


################################
### Código para buscar a cotação do Ethereum para Real


# Moeda que iremos cotar
moedaCotacaoDe = 'ethereum'

# A cotação será de ETH/BRL
moedaCotacaoPara = 'brl'

# Link html que irá retornar a cotação que queremos.
# Você pode colar esse link no seu navegador e você verá que os valores serão retornados na sua tela :)
# link: https://api.coingecko.com/api/v3/simple/price?ids=ethereum&vs_currencies=brl
url = 'https://api.coingecko.com/api/v3/simple/price?ids='+ moedaCotacaoDe +'&vs_currencies=' + moedaCotacaoPara  # URL da API CoinGecko para o preço do Bitcoin em BRL (Real Brasileiro)
    
# Realizar a chamada do link acima.
response = requests.get(url)

# Receber a resposta no format json (Você só precisa saber que json é um padrão de resposta)
data = response.json()

# Retirar o dado da resposta da variável "data" e coloca-la dentro da variavel "bitcoin_price_brl"
ethereum_price_brl = data[moedaCotacaoDe][moedaCotacaoPara]

# Exiba a cotação
print(f'Cotação do ethereum  é: R$ {ethereum_price_brl:.2f}')


################################################################
### Alterar planilha do excel
################################################################

# importar a biblioteca que irá abrir o arquivo excel e realizar a edição
import openpyxl

# Nome do arquivo que iremos abrir e editar.
# Este arquivo deve está na mesma pasta que este código, ok? 
file = 'planilha.xlsx'

# Usando a sequencia de ´codigo que irá carregar o arquivo excel
workbook = openpyxl.load_workbook(file)

# com o arquivo excel aberto, teremos que pegar a aba "Planilha1", 
# certifique que este nome esteja igual ao da sua planilha em excel.
# Caso seu excel esteja em inglês, o nome poderá ser "Sheet1".
sheet = workbook['Planilha1']

# Iremos passar por todas as linhas dessa planilha procurando as textos das cotações
for row in sheet.iter_rows():
    
    # dentro da linha, vamos andar todas as células para procurar os textos das cotações
    for cell in row:
        
        # Se o valor da célula não estiver vazio, iremos realizar a procura que queremos
        if cell.value is not None:
            
            # Pega o valor da célula e transforma ela em texto, pois o que procuramos é um texto.
            value = str(cell.value)
            
            # Verifica se encontramos a célula da "Cotação do dolar"
            if 'Cotação do dolar:' in value:
                
                # Caso tenhamos encontrado está célula
                # vamos pegar o valor da coluna "Neste caso será B"
                column = cell.column_letter
                
                # Vamos pegar o valor da coluna, neste caso é a B e vamos pular 2 colunas para frente "D" para colocarmos o valor da cotação.
                column = chr(ord(column) + 2)
                
                # Vamos pegar o valor da linha.
                row = cell.row
                
                # A célula que receberá o valor da cotação do dolar é D5
                celula = column + str(row)
                
                # Colocar o valor do dolar na célula encontrada.
                sheet[celula] = float(usd_rate)
                
            elif 'Cotação do bitcoin:' in value:
                column = cell.column_letter
                column = chr(ord(column) + 2)
                row = cell.row
                celula = column + str(row)
                sheet[celula] = float(bitcoin_price_brl)
                
            elif 'Cotação do Ethereum:' in value:
                column = cell.column_letter
                column = chr(ord(column) + 2)
                row = cell.row
                celula = column + str(row)
                sheet[column + str(row)] = float(ethereum_price_brl)
                    
    workbook.save(file)
    
print('Atualização realizada com sucesso.')