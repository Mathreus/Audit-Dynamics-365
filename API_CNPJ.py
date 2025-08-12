import requests
import pandas as pd
import time
import certifi
import urllib3

# Suprime apenas o aviso de SSL inseguro
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Carrega os dados garantindo que CNPJ seja lido como texto
df1 = pd.read_excel(
    r'Caminho/Arquivo.xlsx',
    dtype={'CNPJ': str}
)

# Remove caracteres não numéricos e preenche com zeros à esquerda
df1['CNPJ'] = df1['CNPJ'].astype(str).str.replace(r'\D', '', regex=True).str.zfill(14)
cnpjs = df1['CNPJ']

# DataFrame para armazenar os resultados
resultados = pd.DataFrame(columns=['CNPJ', 'SITUACAO'])

def consulta_cnpj(cnpj):
    cnpj_formatado = str(cnpj).zfill(14)
    url = f'https://www.receitaws.com.br/v1/cnpj/{cnpj_formatado}'
    params = {
        "token": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
        "plugin": "RF"
    }
    
    try:
        # Primeira tentativa com SSL verificado
        response = requests.get(url, params=params, timeout=10, verify=certifi.where())
        response.raise_for_status()
        dados = response.json()
        time.sleep(30)
        return dados.get('situacao', 'NÃO ENCONTRADO')

    except requests.exceptions.SSLError:
        # Segunda tentativa ignorando SSL
        response = requests.get(url, params=params, timeout=10, verify=False)
        response.raise_for_status()
        dados = response.json()
        time.sleep(30)
        return dados.get('situacao', 'NÃO ENCONTRADO')

    except Exception:
        return "ERRO NA CONSULTA"

# Processa cada CNPJ
for cnpj in cnpjs:
    situacao = consulta_cnpj(cnpj)
    cnpj_formatado = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
    print(f"{cnpj_formatado} | {situacao}")
    resultados.loc[len(resultados)] = [cnpj_formatado, situacao]

# Salva os resultados
resultados.to_excel('Resultados_Situacao_CNPJ.xlsx', index=False)