import pdfplumber
import re

# Caminho para o arquivo PDF
caminho_pdf = r"C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2025\08. Agosto\Avic 25.2\Folha de Pagamento/Extrato Unimed Campos Gerais.pdf"

# Lista para armazenar os resultados
resultados = []

# Abrir e ler o PDF
with pdfplumber.open(caminho_pdf) as pdf:
    texto_completo = ""
    for pagina in pdf.pages:
        texto_completo += pagina.extract_text() + "\n"

    # Expressão regular para extrair responsável e total família
    padrao = r"Responsável:\s+([A-Z\s]+).*?Total Família:\s*(-?\d{1,3}(?:\.\d{3})*,\d{2})"
    resultados = re.findall(padrao, texto_completo, re.DOTALL)

# Exibir os resultados
for nome, valor in resultados:
    print(f"Responsável: {nome.strip()} | Total Família: R$ {valor}")