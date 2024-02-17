import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import shutil

# Excel
col_nomes = "NOME DOC"
col_senhas = "SENHA"
col_caminho = "DIRETORIO DOC"
arquivo_excel = "./senha/Senhas_pdfs_assoc.xlsm"
#arquivo_excel = "./senha/teste.xlsx"

read = pd.read_excel(arquivo_excel, dtype={col_senhas: str, col_nomes: str})

valor_nome = read[col_nomes].tolist()
valor_senha = read[col_senhas].tolist()
valor_caminho = read[col_caminho].tolist()

dicionario = {id: str(senha) for id, senha in zip(valor_caminho, valor_senha)}

# PDF
index = 0
count = 0
erros = []
mensagens = []
writer = PdfWriter()

# Compressor de PDF
def comprimir_arquivo(arquivo):
    global count
    print(f"\nComprimindo arquivo, aguarde...\nNome: {valor_nome[count]}")
    caminho_app = (r"C:\\Users\\Luccas\\Downloads\\gerador_senha\\pdfsizeopt\\pdfsizeopt")
    arquivo_entrada = arquivo
    arquivo_saida = arquivo
    os.system(f"{caminho_app} --use-pngout=no {arquivo_entrada} {arquivo_saida}")
    print(f"\nCompressão concluída com sucesso!\nArquivo: {valor_nome[count]}")

for pdf in valor_caminho:
    tamanho_pdf = os.path.getsize(pdf)
    pdf_kb = tamanho_pdf / 1024
    if pdf_kb > 800:
        comprimir_arquivo(pdf)
        count += 1

# Gerador de senha PDF
for id, senha in dicionario.items():
    try:
        count = 1
        reader = PdfReader(id)
        if reader.is_encrypted:
            print("Removendo senha...")
            reader.decrypt(senha)
        for page in reader.pages:
            writer.add_page(page)
            print(f"Lendo páginas, aguarde...\n{count}/{len(reader.pages)}")
            count+=1
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
