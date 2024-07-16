import os
import win32com.client
import zipfile
import re
import shutil
import pandas as pd

import win32com.client as win32
import easygui
from datetime import date



mensagem = "Deseja executar o envio de e-mail xxxx?"
titulo = "Confirmação"
opcoes = ["Sim", "Não"]

resposta = easygui.buttonbox(mensagem, title=titulo, choices=opcoes)

if resposta == "Não" or resposta is None:
    easygui.msgbox("A execução foi cancelada pelo usuário", "Execução Cancelada")


else:


    def limpar_diretorio(caminho):
        for root, dirs, files in os.walk(caminho, topdown=False):
            for file in files:
                os.remove(os.path.join(root, file))
            for dir in dirs:
                shutil.rmtree(os.path.join(root, dir))


# Função para salvar anexos e criar pastas
    def salvar_anexos(email, caminho_salvar):
        for anexo in email.Attachments:
            nome_anexo = anexo.FileName
            nome_pasta_match = re.search(r'^([A-Z]+(?:\d+)?)(?:_|\d+)*', nome_anexo)
            nome_pasta = nome_pasta_match.group(1) if nome_pasta_match else 'Outros'

            caminho_pasta = os.path.join(caminho_salvar, nome_pasta)

            if not os.path.exists(caminho_pasta):
                os.makedirs(caminho_pasta)

            caminho_arquivo = os.path.join(caminho_pasta, nome_anexo)
            anexo.SaveAsFile(caminho_arquivo)
            print(f"Anexo '{nome_anexo}' salvo em '{caminho_pasta}'")


    def salvar_anexos_com_data(email, caminho_salvar):
        pasta_data = os.path.join(caminho_salvar, date.today().strftime("%Y-%m-%d"))

        if not os.path.exists(pasta_data):
            os.makedirs(pasta_data)

        for anexo in email.Attachments:
            nome_anexo = anexo.FileName
            caminho_arquivo = os.path.join(pasta_data, nome_anexo)
            anexo.SaveAsFile(caminho_arquivo)
            print(f"Anexo '{nome_anexo}' salvo em '{pasta_data}'")

    # Função para extrair arquivos zipados
    def extrair_arquivos(caminho):
        for root, _, files in os.walk(caminho):
            for file in files:
                if file.endswith(".zip"):
                    caminho_zip = os.path.join(root, file)
                    pasta_destino = os.path.splitext(caminho_zip)[0]
                    with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
                        zip_ref.extractall(pasta_destino)
                    print(f"Arquivos extraídos para '{pasta_destino}'")


    # Função para renomear os arquivos
    def renomear_arquivos(caminho):
        for root, _, files in os.walk(caminho):
            nome_pasta = os.path.basename(root)
            for file in files:
                nome_arquivo, extensao = os.path.splitext(file)
                novo_nome = f"{nome_arquivo} {nome_pasta}{extensao}"
                os.rename(os.path.join(root, file), os.path.join(root, novo_nome))
                print(f"Arquivo renomeado para '{novo_nome}'")




    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    caixa_entrada = outlook.GetDefaultFolder(6)  # Pasta de entrada

    # Campos a serem preenchidos:
    titulo_email = "xxxx"
    caminho_salvar = 'xxx''  # Diretório para salvar os anexos
    caminho_salvar_2 = r'xxxxxx'  # Novo diretório com a data


    # Antes de salvar os novos anexos, vamos limpar o diretório
    limpar_diretorio(caminho_salvar)

    # Iterar pelos e-mails
    for email in reversed(list(caixa_entrada.Items)):
        if titulo_email in email.Subject:
            salvar_anexos_com_data(email, caminho_salvar_2)
            salvar_anexos(email, caminho_salvar)
            extrair_arquivos(caminho_salvar)
            renomear_arquivos(caminho_salvar)
            break


    ################# CODIGO DE ENVIO

    # Diretório onde os arquivos foram extraídos
    caminho_arquivos = r'C:xxxxxxx'

    # Função para enviar e-mail com os arquivos correspondentes a uma cooperativa
    def enviar_arquivos_cooperativa(cooperativa, destinatario):
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.Subject = f'{cooperativa} xxxxxx '
        email.To = destinatario
        email.Body = f'Prezado(s), segue(em) anexo(s) o(s) arquivo(s) com as informações do xxxxxxx correspondente(s) à cooperativa {cooperativa}.'

        # Anexar os arquivos ao e-mail
        for root, _, files in os.walk(caminho_arquivos):
            for file in files:
                cooperativa_arquivo = re.match(r'^(\d{4})', file)
                if cooperativa_arquivo and cooperativa_arquivo.group(1) == cooperativa:
                    arquivo = os.path.join(root, file)
                    email.Attachments.Add(arquivo)

        # Enviar o e-mail
        try:
            email.Send()
            print(f'E-mail enviado para {destinatario} com os arquivos da cooperativa {cooperativa}')
        except Exception as e:
            print(f'Erro ao enviar e-mail para {destinatario}: {str(e)}')

    # Carregar a planilha
    caminho_planilha = r'C:\xxxxx.xlsx'
    tabela_emails = pd.read_excel(caminho_planilha)




    for index, row in tabela_emails.iterrows():
        cooperativa = str(row['cooperativa'])[:4]  # Pegar os 4 primeiros dígitos
        destinatario = row['e-mail']
        enviar_arquivos_cooperativa(cooperativa, destinatario)

    print("executado com sucesso!")



