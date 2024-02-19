import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

index = 0
opcao_comprimir = False

# Manipulação do excel
col_nomes = "NOME DOC"
col_senhas = "SENHA"
col_caminho = "DIRETORIO DOC"
arquivo_excel = "./planilhas_senha/senha_ass_serv(teste).xlsm"

read = pd.read_excel(arquivo_excel, dtype={col_senhas: str, col_nomes: str})

valor_nome = read[col_nomes].tolist()
valor_senha = read[col_senhas].tolist()
valor_caminho = read[col_caminho].tolist()

dicionario = {id: str(senha) for id, senha in zip(valor_caminho, valor_senha)}

# Compressor de PDF
def comprimir_arquivo(arquivo, index):
    tamanho_pdf = os.path.getsize(arquivo)
    pdf_kb = tamanho_pdf / 1024
    if pdf_kb > 1100:
        print(f"\nComprimindo arquivo, aguarde...\nNome: {valor_nome[index]}")
        caminho_app = (r"C:\\Users\\Lucas.Aguiar\\Desktop\\gerador_senha\\otimizador_pdf\\executavel")
        arquivo_entrada = arquivo
        arquivo_saida = arquivo
        os.system(f"{caminho_app} --use-pngout=no {arquivo_entrada} {arquivo_saida}")
        print(f"\nCompressão concluída com sucesso!\nArquivo: {valor_nome[index]}")

# Gerador de senha PDF
def gerador_senha(id, senha, index):
    global falhas
    count = 1

    try:
        with open(id, "rb") as arquivo_entrada:
            reader = PdfReader(arquivo_entrada)
            writer = PdfWriter()
            print(f"\nAdicionando senha ao PDF \"{valor_nome[index]}\", aguarde...")
            for page in reader.pages:
                writer.add_page(page)
                if not reader.is_encrypted:
                    print(f"Página: {count}/{len(reader.pages)}")
                    writer.encrypt(user_password=senha)
                    count+=1
            with open(id, "wb") as arquivo_saida:
                writer.write(arquivo_saida)
            print(f"\nArquivo: {valor_nome[index]}\nSenha: {senha}\nPosição: {index+1}/{len(dicionario)}\n")
    except Exception as e:
        print(f"Erro no arquivo: {valor_nome[index]}\nErro: {e}\nPosição: {index+1}/{len(dicionario)}\n")
        falhas.append(e)

# Executar compressor de PDF
if opcao_comprimir == True:
    count = 0
    for pdf in valor_caminho:
        comprimir_arquivo(pdf, count)
        count += 1
    print("\nCompressão executada com sucesso!")

# Executar gerador de senha para PDF
falhas = []
for id, senha in dicionario.items():
    gerador_senha(id, senha, index)
    index += 1
    if index > len(dicionario):
        print("\nProcesso concluido!\n")
        if falhas > 0:
            print(f"Houveram {len(falhas)} erros durante o processo! As seguintes falhas foram encontradas:\n")
            for erro in falhas:
                print(erro + "\n")
        else:
            print(f"Total de arquivos processados: {len(dicionario)}")
    


