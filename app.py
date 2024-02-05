import pandas as pd
from PyPDF2 import PdfWriter

# Excel
col_nomes = "NOME ARQUIVO"
col_senhas = "SENHA ARQUIVO"
arquivo_excel = "./senha/senhas_condominos.xlsx"

read = pd.read_excel(arquivo_excel)
valor_nome = read[col_nomes].tolist()
valor_senha = read[col_senhas].tolist()

dicionario = {id: str(senha) for id, senha in zip(valor_nome, valor_senha)}
for id, senha in dicionario.items():
   print(f"Nome: {id}\nSenha: {senha}")

# PDF
writer = PdfWriter()

for id, senha in dicionario.items():
    writer.encrypt(senha)
    print(f"Condômino: {id}\nSenha: {senha}")

