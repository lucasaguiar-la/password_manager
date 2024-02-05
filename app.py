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
            reader.decrypt(senha)
        for page in reader.pages:
            writer.add_page(page)
            if reader.is_encrypted == False:
                writer.encrypt(user_password=senha)
            with open(id, "wb") as f:
                writer.write(f)
        print(f"Arquivo: {valor_nome[index]}\nSenha: {senha}\nPosição: {index+1}/{len(dicionario)}\n")
    except Exception as e:
        print(f"Erro no arquivo: {valor_nome[index]}\nErro: {e}\nPosição: {index+1}/{len(dicionario)}\n")
        erros.append(valor_nome[index])
        mensagens.append(e)

    index += 1

print(f"Arquivos com senha: {len(dicionario) - len(erros)}\nFalhas: {len(erros)}")
print(f"\nArquivos com falhas: {erros}\n")
index = 0
print("=" * 30 + "\n" + "======== [LOG DE ERROS] ======\n" + "=" * 30)

for erro in erros:
    print(f"Arquivo: {erros[index]}\nMensagens: {mensagens[index]}\n")
