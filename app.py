import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

# Excel
col_caminho = "CAMINHO ARQUIVO"
col_senhas = "SENHA ARQUIVO"
col_nomes = "NOME ARQUIVO"
arquivo_excel = "./senha/senhas_condominos.xlsx"

read = pd.read_excel(arquivo_excel)
valor_caminho = read[col_caminho].tolist()
valor_senha = read[col_senhas].tolist()
valor_nome = read[col_nomes].tolist()

dicionario = {id: str(senha) for id, senha in zip(valor_caminho, valor_senha)}

# PDF
index = 0
writer = PdfWriter()

for id, senha in dicionario.items():
    reader = PdfReader(id)
    if reader.is_encrypted:
        reader.decrypt(senha)
    for page in reader.pages:
        writer.add_page(page)
        if reader.is_encrypted == False:
            writer.encrypt(user_password=senha)
        with open(id, "wb") as f:
            writer.write(f)
    print(f"Arquivo: {valor_nome[index]}\nSenha: {senha}")
    index += 1

