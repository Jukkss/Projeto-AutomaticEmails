# Projeto de aprendizado com Python - AutomaticEmails
import email
import imaplib
import re
import os
from dotenv import load_dotenv

# Carrega variáveis do arquivo .env
load_dotenv()

EMAIL = os.getenv("GMAIL_USER")
PASSWORD = os.getenv("GMAIL_PASSWORD")
SERVER = 'imap.gmail.com'

def ler_email_do_gmail():
    try:
        mail = imaplib.IMAP4_SSL(SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select('inbox')

        result, data = mail.search(None, 'ALL')
        mail_ids = []

        for block in data:
            mail_ids += block.split()

        for i in mail_ids:
            status, data = mail.fetch(i, '(RFC822)')

            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    mail_from = msg['From']
                    mail_subject = msg['Subject']

                    mail_content = ''
                    if msg.is_multipart():
                        for part in msg.get_payload():
                            if part.get_content_type() == 'text/plain':
                                mail_content += part.get_payload(decode=True).decode()
                    else:
                        mail_content = msg.get_payload(decode=True).decode()

                    tipodDeDado = mail_subject
                    mensagem = mail_content

                    if 'Resultado de Avaliação da Disciplina' in tipodDeDado:
                        try:
                            match = re.search(r'Prezado\(a\).*', mensagem, re.DOTALL)
                            mensagem_limpa = match.group(0) if match else mensagem

                            blocos = re.split(r'(AVALIAÇÃO DO\(A\) PROFESSOR\(A\)|AVALIAÇÃO DA DISCIPLINA|COORDENAÇÃO DO CURSO|CONDIÇÕES INSTITUCIONAIS|SISTEMA CANVAS|SISTEMA MICROSOFT TEAMS)', mensagem_limpa)

                            resultados = {}

                            for i in range(1, len(blocos), 2):
                                nome_bloco = blocos[i]
                                conteudo_bloco = blocos[i + 1]

                                media_match = re.search(r'Média das Avaliações\s+([\d,.]+)', conteudo_bloco)
                                if media_match:
                                    media = media_match.group(1)
                                    resultados[nome_bloco] = media

                            print(f"Resultados para '{tipodDeDado}': {resultados}")

                        except Exception as e:
                            print(f"Erro ao processar o email '{tipodDeDado}': {e}")

        mail.logout()

    except Exception as e:
        print(f"Erro geral: {e}")