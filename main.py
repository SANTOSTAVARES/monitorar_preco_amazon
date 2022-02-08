import xlsxwriter
import requests
from bs4 import BeautifulSoup

def main():
    """Realiza uma busca no site da 'Amazon.com.br' pela palavra 'Iphone'.
    E registra num arquivo excel, os resultados dos nomes e preços dos produtos."""

    url_amazon = "https://www.amazon.com.br/s?k=iphone&__mk_pt_BR=%C3%85M%C3%85%C5%BD%C3%95%C3%91&ref=nb_sb_noss" #URL com os resultados da busca pela palavra Iphone no site da Amazon.
    navegador = {
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36"
    } #Informações a serem enviadas, para evitar que o site recuse a resposta HTTP.
    
    try:
        #Se o link não estiver disponível na rede, irá apresentar um erro.
        requisicao_codigo = requests.get(url_amazon, headers=navegador) 
    except Exception as erro:
        print(f'Ao realizar a requisição HTTP, ocorreu o seguinte erro:\n {erro}')
    else:    
        conteudo_html = BeautifulSoup(requisicao_codigo.content, 'html.parser')
        SKU = conteudo_html.find_all('div', class_="sg-col-inner")
        
        #if conteudo_html != None: #Valida se a resposta HTTP foi realizada como vazia.
                        
        with xlsxwriter.Workbook('cotacao.xlsx') as planilha: #Cria arquivo Excel.
            contador = 1
            nova_aba = planilha.add_worksheet()
            
            #Cria cabeçalho para o registro dos dados.
            nova_aba.write('A1', 'Nome')
            nova_aba.write('B1', 'Preço_R$')
            
            for x in SKU:
                if x.find('span', class_="a-size-base-plus a-color-base a-text-normal"):
                    '''Registra o nome de um produto na planilha.'''
                    nome_produto = x.find('span', class_="a-size-base-plus a-color-base a-text-normal").get_text().strip()
                    nova_aba.write(contador, 0, nome_produto)
                    contador += 1  
                    
                if x.find('span', class_="a-offscreen"):
                    '''Registra o preço de um produto na planilha.'''
                    preco_produto = x.find('span', class_="a-offscreen").get_text().strip()
                    nova_aba.write(contador - 1, 1, preco_produto[2:]) #A seleção da linha do excel como "contador - 1", é para ajuste do registro na linha correta.
                        
        if contador > 1: #O contador é a referência se foram encontrados Tags HTML com nome de produtos. Caso não tenha sido encontrado, é por falta de resposta da requisição HTTP.
            print('A planilha está salva com sucesso.\nAlguns produtos podem terem sido registrados sem a informação do preço.')
        else:
            print('Não foi possível consultar o HTML do site por meio da requisição HTTP. Por favor, tente novamente.')

if __name__ == "__main__":
    main()