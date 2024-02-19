import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

# Manipulação do excel
col_nomes = "NOME DOC"
col_senhas = "SENHA"
col_caminho = "DIRETORIO DOC"
arquivo_excel = "./senha/Senhas_pdfs_assoc(gui).xlsm"

read = pd.read_excel(arquivo_excel, dtype={col_senhas: str, col_nomes: str})

valor_nome = read[col_nomes].tolist()
valor_senha = read[col_senhas].tolist()
valor_caminho = read[col_caminho].tolist()

dicionario = {id: str(senha) for id, senha in zip(valor_caminho, valor_senha)}

# Compressor de PDF
def comprimir_arquivo(arquivo):
    count = 1
    print(f"\nComprimindo arquivo, aguarde...\nNome: {valor_nome[count]}")
    caminho_app = (r"./pdfsizeopt/pdfsizeopt")
    arquivo_entrada = arquivo
    arquivo_saida = arquivo
    os.system(f"{caminho_app} --use-pngout=no {arquivo_entrada} {arquivo_saida}")
    print(f"\nCompressão concluída com sucesso!\nArquivo: {valor_nome[count]}")

# Gerador de senha PDF
def gerador_senha(id, senha):
    index = 0
    erros = []
    mensagens = []
    count = 1

    try:
        with open(id, "rb") as arquivo_entrada:
            reader = PdfReader(arquivo_entrada)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
                print("\nAdicionando senha, aguarde...")
                if not reader.is_encrypted:
                    print(f"Página: {count}/{len(reader.pages)}")
                    writer.encrypt(user_password=senha)
                    count+=1
            with open(id, "wb") as arquivo_saida:
                writer.write(arquivo_saida)
            print(f"\nArquivo: {valor_nome[index]}\nSenha: {senha}\nPosição: {index+1}/{len(dicionario)}\n")
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

# Executar compressor de PDF
'''
for pdf in valor_caminho:
    count = 0
    tamanho_pdf = os.path.getsize(pdf)
    pdf_kb = tamanho_pdf / 1024
    if pdf_kb > 1000:
        comprimir_arquivo(pdf)
        count += 1
'''

# Executar gerador de senha para PDF
for id, senha in dicionario.items():
    gerador_senha(id, senha)
