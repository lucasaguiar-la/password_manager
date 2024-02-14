import os
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
erros = []
mensagens = []
writer = PdfWriter()

for id, senha in dicionario.items():
    try:
        reader = PdfReader(id)
        if reader.is_encrypted:
            print("Removendo senha...")
            reader.decrypt(senha)
        for page in reader.pages:
            writer.add_page(page)
            if reader.is_encrypted == False:
                print("Criando senha...")
                writer.encrypt(user_password=senha)
            with open(id, "wb") as f:
                writer.write(f)
        print(f"Arquivo: {valor_nome[index]}\nSenha: {senha}\nPosição: {index+1}/{len(dicionario)}\n")
    except Exception as e:
        print(f"Erro no arquivo: {valor_nome[index]}\nErro: {e}\nPosição: {index+1}/{len(dicionario)}\n")
        erros.append(valor_nome[index])
        mensagens.append(e)
    index += 1

print(f"Arquivos com sucesso: {len(dicionario) - len(erros)}\nArquivos com falhas: {len(erros)}")

# Log de erros
if len(erros) > 0:
    index = 0
    print(f"\nArquivos com falhas: {erros}\n")
    print("=" * 30 + "\n" + "======== [LOG DE ERROS] ======\n" + "=" * 30 + "\n")

    for erro in erros:
        print(f"Arquivo: {erro}\nMensagens: {mensagens[index]}\n")
        index+1
