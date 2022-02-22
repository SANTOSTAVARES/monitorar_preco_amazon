import xlsxwriter
import requests
from bs4 import BeautifulSoup
import pyautogui
import pyperclip
import clipboard
import time 

def main():
    """Realiza uma busca no site da 'Amazon.com.br' pela palavra 'Iphone'.
    E registra num arquivo excel, os resultados dos nomes e preços dos produtos."""

    pyautogui.hotkey('win', 'x') #Abre o menu rápido de aplicativos.
    time.sleep(3)
    pyautogui.click(x=1482, y=416) #Clica no PowerShell
    time.sleep(10)
    pyautogui.write('start chrome www.amazon.com.br') #No PowerShell roda o comando que já abre o site da Amazon direto no Google Chrome
    pyautogui.press('enter')
    time.sleep(8)
    pyautogui.click(x=2000, y=110) #Clica no campo de pesquisa para procurar por Iphone
    pyautogui.write('iphone')
    pyautogui.press('enter')
    pyautogui.click(x=2000, y=50)
    time.sleep(10)
    pyautogui.hotkey('ctrl', 'c') #Copia o link da pagina que será útil para requisição HTTP a ser feita.
    
    url_amazon = clipboard.paste() #URL com os resultados da busca pela palavra Iphone no site da Amazon.
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
