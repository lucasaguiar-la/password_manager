import os 

def comprimir_arquivo(arquivo):
    os.system(f"C:\\Users\\Lucas.Aguiar\\Desktop\\gerador_senha\\pdfsizeopt\\pdfsizeopt.exe C:\\Users\\Lucas.Aguiar\\Desktop\\gerador_senha\\pdf\{arquivo} C:\\Users\\Lucas.Aguiar\\Desktop\\gerador_senha\\pdfsizeopt\\pdf\\{arquivo}")

nome_arquivo = "teste.pdf"
comprimir_arquivo(nome_arquivo)