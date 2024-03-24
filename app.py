import requests
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO
import time  

dados = pd.read_excel('tags.xlsx')
tags = dados.iloc[:, 0]  

headers = {
    'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
}

# Criar um objeto ExcelWriter
with pd.ExcelWriter('dados_coletados.xlsx') as writer:
    for tag in tags:  
        url = f'https://www.fundamentus.com.br/fii_proventos.php?papel={tag}&tipo=2'
        site = requests.get(url, headers=headers)
        
        if site.status_code == 200:
            content = site.content
            soup = BeautifulSoup(content, 'html.parser')
            tabela = soup.find(name='table')
            
            if tabela:  
                tabela_string = str(tabela)
                df = pd.read_html(StringIO(tabela_string))[0]  
                
                # Ordenar os dados pela primeira coluna
                df = df.sort_values(by=df.columns[0])  
                
                # Formatar os valores para substituir ',' por '.'
                df = df.map(lambda x: str(x).replace(',', '.') if isinstance(x, (int, float)) else x)
                
                df['Tag'] = tag
                df.to_excel(writer, sheet_name='Dados', index=False, startrow=(tags.tolist().index(tag) * (len(df) + 2)) + 1)
                print(f"Coletado de {tag}")
            else:
                print(f"Nenhuma tabela encontrada para a tag {tag}")
        else:
            print(f"Erro ao acessar a página para a tag {tag}")
        
        time.sleep(1)

print("Dados salvos na planilha 'Dados' do arquivo dados_coletados.xlsx")
