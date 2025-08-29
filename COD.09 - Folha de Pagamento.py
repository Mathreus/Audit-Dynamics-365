import re
import pdfplumber
import difflib
import pandas as pd
import unicodedata

# Função para normalizar nomes
def normalizar_nome(nome):
    nome = unicodedata.normalize('NFKD', nome)
    nome = ''.join(c for c in nome if not unicodedata.combining(c))
    return nome.upper().strip()

# Função para extrair valores do TXT (Unimed + Odonto)
def extrair_valores_txt(caminho_txt):
    with open(caminho_txt, "r", encoding="latin1") as arquivo:
        texto = arquivo.read()

    colaboradores = re.findall(r"(\d{6})\s+([A-ZÀ-Ú\s\-]+)", texto)
    valores_txt = {}

    for codigo, nome in colaboradores:
        nome_limpo = normalizar_nome(' '.join(nome.strip().split()))
        padrao_bloco = rf"{codigo}\s+{re.escape(nome.strip())}(.*?)(?=\d{{6}}\s+[A-ZÀ-Ú\s\-]+|\Z)"
        bloco_match = re.search(padrao_bloco, texto, re.DOTALL)
        bloco = bloco_match.group(1) if bloco_match else ""

        titular = re.search(r"Unimed.*?- T.*?([\d.,]+)-", bloco)
        dependente = re.search(r"Unimed.*?- D.*?([\d.,]+)-", bloco)
        odontos = re.findall(r"Dental Uni Odonto.*?([\d.,]+)-", bloco)

        valor_t = float(titular.group(1).replace('.', '').replace(',', '.')) if titular else 0.0
        valor_d = float(dependente.group(1).replace('.', '').replace(',', '.')) if dependente else 0.0
        valor_odontologico = sum(float(v.replace('.', '').replace(',', '.')) for v in odontos)

        valores_txt[nome_limpo] = {
            "valor_unimed": round(valor_t + valor_d, 2),
            "valor_odontologico": round(valor_odontologico, 2)
        }

    return valores_txt

# Função para extrair valores do PDF
def extrair_valores_pdf(caminho_pdf):
    resultados = {}
    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + "\n"

        padrao = r"Responsável:\s+([A-ZÀ-Ú\s]+).*?Total Família:\s*(-?\d{1,3}(?:\.\d{3})*,\d{2})"
        matches = re.findall(padrao, texto_completo, re.DOTALL)

        for nome, valor in matches:
            nome_limpo = normalizar_nome(' '.join(nome.strip().split()))
            valor_float = float(valor.replace('.', '').replace(',', '.').replace('-', '')) if valor else 0.0
            resultados[nome_limpo] = round(valor_float, 2)

    return resultados

# Função para encontrar nome mais próximo
def encontrar_nome_proximo(nome, lista_nomes, threshold=0.8):
    similares = difflib.get_close_matches(nome, lista_nomes, n=1, cutoff=threshold)
    return similares[0] if similares else None

# Caminhos dos arquivos
caminho_txt = r"C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2025\08. Agosto\Avic 25.2\Folha de Pagamento\Folha de Pagamento - Campos Gerais.txt"
caminho_pdf = r"C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2025\08. Agosto\Avic 25.2\Folha de Pagamento\Extrato Unimed Campos Gerais.pdf"

# Extrair dados
valores_txt = extrair_valores_txt(caminho_txt)
valores_pdf = extrair_valores_pdf(caminho_pdf)

# Criar lista de registros para tabela
registros = []

for responsavel_pdf, valor_pdf in valores_pdf.items():
    nome_txt = encontrar_nome_proximo(responsavel_pdf, valores_txt.keys())
    metade_extrato = round(valor_pdf / 2, 2)

    if nome_txt:
        valor_unimed = valores_txt[nome_txt]["valor_unimed"]
        valor_odontologico = valores_txt[nome_txt]["valor_odontologico"]
        diferenca = round(valor_unimed - metade_extrato, 2)

        registros.append({
            "Responsável (PDF)": responsavel_pdf.title(),
            "Nome Correspondente (TXT)": nome_txt.title(),
            "Valor Descontado na Folha (Unimed)": valor_unimed,
            "Valor Total Família (Extrato)": valor_pdf,
            "Metade do Valor do Extrato": metade_extrato,
            "Diferença (Desconto - Metade Extrato)": diferenca,
            "Valor Odontológico (TXT)": valor_odontologico  # Agora como última coluna
        })
    else:
        registros.append({
            "Responsável (PDF)": responsavel_pdf.title(),
            "Nome Correspondente (TXT)": "Não encontrado",
            "Valor Descontado na Folha (Unimed)": None,
            "Valor Total Família (Extrato)": valor_pdf,
            "Metade do Valor do Extrato": metade_extrato,
            "Diferença (Desconto - Metade Extrato)": None,
            "Valor Odontológico (TXT)": None  # Última coluna
        })

# Criar DataFrame e exportar para Excel
df = pd.DataFrame(registros)

# Reordenar colunas para garantir que odontológico fique por último
colunas_ordenadas = [
    "Responsável (PDF)",
    "Nome Correspondente (TXT)",
    "Valor Descontado na Folha (Unimed)",
    "Valor Total Família (Extrato)",
    "Metade do Valor do Extrato",
    "Diferença (Desconto - Metade Extrato)",
    "Valor Odontológico (TXT)"
]
df = df[colunas_ordenadas]

caminho_saida = r"C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Automações\Folha\Teste\comparativo_unimed_v5.xlsx"
df.to_excel(caminho_saida, index=False)

print(f"✅ Comparativo gerado com sucesso e exportado para:\n{caminho_saida}")
