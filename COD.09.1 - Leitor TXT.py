import re

# Abrir e ler o conteúdo do arquivo
with open(r"C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2025\08. Agosto\Avic 25.2\Folha de Pagamento\Folha de Pagamento - Campos Gerais.txt", "r", encoding="latin1") as arquivo:
    texto = arquivo.read()

# Separar os blocos de colaboradores
blocos = re.split(r"\d{4}\s+[A-ZÀ-Úa-zà-ú\s\/\-]+?\s+\d{6}\s+[A-ZÀ-Ú\s\-]+", texto)

# Encontrar todos os colaboradores (código e nome)
colaboradores = re.findall(r"(\d{6})\s+([A-ZÀ-Ú\s\-]+)", texto)

# Dicionário para armazenar os valores
planos_saude = {}

# Iterar sobre os blocos e colaboradores
for i, (codigo, nome) in enumerate(colaboradores):
    bloco = blocos[i + 1]  # pular o primeiro bloco que é cabeçalho

    # Buscar valores de Unimed - T e Unimed - D (captura o terceiro número com sinal negativo)
    titular = re.search(r"Unimed.*?- T.*?([\d.,]+)-", bloco)
    dependente = re.search(r"Unimed.*?- D.*?([\d.,]+)-", bloco)

    # Converter valores para float e somar
    valor_t = float(titular.group(1).replace('.', '').replace(',', '.')) if titular else 0.0
    valor_d = float(dependente.group(1).replace('.', '').replace(',', '.')) if dependente else 0.0
    total = valor_t + valor_d

    # Limpar o nome removendo o "A" final, se presente
    nome_limpo = ' '.join(nome.strip().split())
    if nome_limpo.endswith("A"):
        nome_limpo = ' '.join(nome_limpo.split()[:-1])

    planos_saude[nome_limpo] = total

# Exibir os resultados
for nome, valor in planos_saude.items():
    print(f"{nome}: R$ {valor:.2f}")