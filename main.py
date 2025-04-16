import re
import os
from imap_tools import MailBox
import pandas as pd
from dotenv import load_dotenv

# Carregando login env
load_dotenv()

# Configurações para Gmail e Outlook
EMAIL_CONFIG = {
    'gmail': {
        'host': 'imap.gmail.com',
        'email': os.getenv('EMAIL_GMAIL'),
        'senha': os.getenv('EMAIL_GMAIL_APP_PASSWORD')
    },
    'outlook': {
        'host': 'imap-mail.outlook.com',
        'email': 'seu-email@outlook.com',
        'senha': 'sua-senha'
    }
}


def extrair_dados_email(corpo):
    dados = {}

    prof_match = re.search(r'AVALIAÇÃO DO\(A\) PROFESSOR\(A\) - (.+?) - CH', corpo)
    media_prof_match = re.search(r'Média das Avaliações\s+(\d\.\d+)', corpo)
    disciplina_match = re.search(r'disciplina (.+?) - CH', corpo)
    curso_match = re.search(r'Curso (.+?), Oferta', corpo)

    if prof_match:
        dados['professor'] = prof_match.group(1).strip()
    if media_prof_match:
        dados['media_professor'] = float(media_prof_match.group(1))
    if disciplina_match:
        dados['disciplina'] = disciplina_match.group(1).strip()
    if curso_match:
        dados['curso'] = curso_match.group(1).strip()

    return dados


def ler_emails(provedor):
    config = EMAIL_CONFIG[provedor]
    resultados = []

    with MailBox(config['host']).login(config['email'], config['senha']) as mailbox:
        for msg in mailbox.fetch(subject='Resultado de Avaliação da Disciplina', bulk=True):
            corpo = msg.text or msg.html or ''
            dados = extrair_dados_email(corpo)
            if dados:
                resultados.append(dados)

    return resultados


if __name__ == '__main__':
    todos_resultados = []

    for prov in EMAIL_CONFIG:
        print(f"Lendo e-mails de: {prov}")
        resultados = ler_emails(prov)
        todos_resultados.extend(resultados)

    df = pd.DataFrame(todos_resultados)
    df.to_csv('relatorio_avaliacoes.csv', index=False)
    print("Relatório gerado: relatorio_avaliacoes.csv")