import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

# Excel
col_nomes = "NOME ARQUIVO"
col_senhas = "SENHA ARQUIVO"
arquivo_excel = "./senha/senhas_condominos.xlsx"

read = pd.read_excel(arquivo_excel)
valor_nome = read[col_nomes].tolist()
valor_senha = read[col_senhas].tolist()

dicionario = {id: str(senha) for id, senha in zip(valor_nome, valor_senha)}

# PDF
writer = PdfWriter()



for id, senha in dicionario.items():
    reader = PdfReader(id)
    for page in reader.pages:
        writer.add_page(page)
        writer.encrypt(user_password=senha)
        with open(id, "wb") as f:
            writer.write(f)
    print(f"Condômino: {id}\nSenha: {senha}")

