import oracledb, openpyxl, smtplib, os, logging
from logging.handlers import RotatingFileHandler
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

# Definindo variáveis
# Consultas:

uis = """
select Data, UIS_TOT from (
    select to_char(trunc(E3TIMESTAMP, 'hh24'), 'dd/mm/yyyy hh24:mi') as Data, 
           round(AVG(US_SCT_P_TOT), 2) as UIS_TOT 
    from HT_UIS_GL 
    where to_char(E3TIMESTAMP, 'MMYYYY') = to_char(SYSDATE - INTERVAL '1' HOUR, 'MMYYYY') 
      and trunc(E3TIMESTAMP, 'hh24') = trunc(SYSDATE - interval '1' hour, 'hh24') 
    group by trunc(E3TIMESTAMP, 'hh24')
)
"""
uso = """
select Data, USO_TOT from (
    select to_char(trunc(E3TIMESTAMP, 'hh24'), 'dd/mm/yyyy hh24:mi') as Data, 
           round(AVG(US_SCT_P_TOT),2) as USO_TOT 
    from HT_USO_GL 
    where to_char(E3TIMESTAMP, 'MMYYYY') = to_char(SYSDATE - INTERVAL '1' HOUR, 'MMYYYY') 
      and trunc(E3TIMESTAMP, 'hh24') = trunc(SYSDATE - interval '1' hour, 'hh24') 
    group by trunc(E3TIMESTAMP, 'hh24')
)
"""
ucr_u1 = """
select Data, round(UCR_U1_MW, 2) as UCR_U1_MW from (
    select to_char(trunc(E3TIMESTAMP, 'hh24'), 'dd/mm/yyyy hh24:mi') as Data, 
           round(AVG(GE_SCT_P_TOT) / 1000, 2) as UCR_U1_MW 
    from HT_UCR_U1
    where to_char(E3TIMESTAMP, 'MMYYYY') = to_char(SYSDATE - INTERVAL '1' HOUR, 'MMYYYY') 
      and trunc(E3TIMESTAMP, 'hh24') = trunc(SYSDATE - interval '1' hour, 'hh24') 
    group by trunc(E3TIMESTAMP, 'hh24')
)

"""
ucr_u2 = """
select Data, round(UCR_U2_MW, 2) as UCR_U2_MW from (
    select to_char(trunc(E3TIMESTAMP, 'hh24'), 'dd/mm/yyyy hh24:mi') as Data, 
           round(AVG(GE_SCT_P_TOT) / 1000, 2) as UCR_U2_MW 
    from HT_UCR_U2
    where to_char(E3TIMESTAMP, 'MMYYYY') = to_char(SYSDATE - INTERVAL '1' HOUR, 'MMYYYY') 
      and trunc(E3TIMESTAMP, 'hh24') = trunc(SYSDATE - interval '1' hour, 'hh24') 
    group by trunc(E3TIMESTAMP, 'hh24')
)
"""
ucr_u3 = """
select Data, round(UCR_U3_MW, 2) as UCR_U3_MW from (
    select to_char(trunc(E3TIMESTAMP, 'hh24'), 'dd/mm/yyyy hh24:mi') as Data, 
           round(AVG(GE_SCT_P_TOT) / 1000, 2) as UCR_U3_MW 
    from HT_UCR_U3
    where to_char(E3TIMESTAMP, 'MMYYYY') = to_char(SYSDATE - INTERVAL '1' HOUR, 'MMYYYY') 
      and trunc(E3TIMESTAMP, 'hh24') = trunc(SYSDATE - interval '1' hour, 'hh24') 
    group by trunc(E3TIMESTAMP, 'hh24')
)
"""
ucr_u4 = """
select Data, round(UCR_U4_MW, 2) as UCR_U4_MW from (
    select to_char(trunc(E3TIMESTAMP, 'hh24'), 'dd/mm/yyyy hh24:mi') as Data, 
           round(AVG(GE_SCT_P_TOT) / 1000, 2) as UCR_U4_MW 
    from HT_UCR_U4
    where to_char(E3TIMESTAMP, 'MMYYYY') = to_char(SYSDATE - INTERVAL '1' HOUR, 'MMYYYY') 
      and trunc(E3TIMESTAMP, 'hh24') = trunc(SYSDATE - interval '1' hour, 'hh24') 
    group by trunc(E3TIMESTAMP, 'hh24')
)
"""
urc = """
select Data, URC_TOT from (
    select to_char(trunc(E3TIMESTAMP, 'hh24'), 'dd/mm/yyyy hh24:mi') as Data, 
           round(AVG(US_SCT_P_TOT),2) as URC_TOT 
    from HT_URC_GL
    where to_char(E3TIMESTAMP, 'MMYYYY') = to_char(SYSDATE - INTERVAL '1' HOUR, 'MMYYYY') 
      and trunc(E3TIMESTAMP, 'hh24') = trunc(SYSDATE - interval '1' hour, 'hh24') 
    group by trunc(E3TIMESTAMP, 'hh24')
)
"""

previous_time = datetime.now() - timedelta(hours=1)

file_date = previous_time.strftime("%d_%m_%Y - %H_00")
excel_date = previous_time.strftime("%d/%m/%Y  %H:00")

# Configurando o log com rotação por número de arquivos
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_handler = RotatingFileHandler(r'K:\Geracao Comum\DVOP\Projetos\Inovvo\Histórico de erros.log', maxBytes=1e6, backupCount=5)
log_handler.setFormatter(log_formatter)
logger = logging.getLogger(__name__)
logger.addHandler(log_handler)
logger.setLevel(logging.DEBUG)


# Definindo Funções

def query(sql_query):
    conn = oracledb.connect(user='cog', password='ro5iww8b', dsn='jira ')
    try:
        cursor = conn.cursor()
        cursor.execute(sql_query)
        result = cursor.fetchone()
        power = float(result[1])
    finally:
        cursor.close()
        conn.close()
    return power


def save_query():

    # Criar um novo arquivo Excel e selecionar a primeira planilha
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    column_a = [83669000, 73694000, 83740000, 90000050]
    column_b = [excel_date] * len(column_a)

    # Preencher as colunas A e B com os valores
    for idx, (value_a, value_b) in enumerate(zip(column_a, column_b)):
        sheet.cell(row=idx + 1, column=1, value=value_a)
        sheet.cell(row=idx + 1, column=2, value=value_b)

    # Preencher coluna C com valores das potências de cada usina
    sheet['C1'] = query(urc)
    sheet['C2'] = query(ucr_u1) + query(ucr_u2)
    sheet['C3'] = query(uso)
    sheet['C4'] = query(uis)
    sheet['D2'] = query(ucr_u3) + query(ucr_u4)

    # Ajustar o tamanho da coluna B
    sheet.column_dimensions['B'].width = 16

    # Salvar o arquivo Excel com o nome fornecido
    workbook.save(os.path.join(fr"K:\Geracao Comum\DVOP\Projetos\Inovvo\Relatórios", f"Relatório Usinas - {file_date}.xlsx"))


def send_query():
    # Configurações do servidor SMTP do Outlook
    smtp_host = 'smtp.celesc.com.br'
    smtp_port = 25
    username = 'cog@celesc.com.br'
    receiver = 'dados@inovvodata.com.br'

    # Criando o objeto MIMEMultipart para compor o e-mail
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = receiver
    msg['Subject'] = 'Relatório - Usinas CELESC'

    email_body = ""

    # Adicionando o corpo do e-mail
    msg.attach(MIMEText(email_body, 'plain'))

    # Caminho absoluto para o arquivo que será anexado ao e-mail
    file_path = fr"K:\Geracao Comum\DVOP\Projetos\Inovvo\Relatórios\Relatório Usinas - {file_date}.xlsx"

    # Adicionando o caminho absoluto à lista de arquivos anexos
    files = [file_path]
    for file in files:
        with open(file, 'rb') as attachment:
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet', name=f"Relatório Usinas - {file_date}.xlsx")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment')
            msg.attach(part)

    # Conectando ao servidor SMTP e enviando o e-mail
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.set_debuglevel(1)
            server.sendmail(username, msg['To'], msg.as_string())
            server.quit()
 

# Chamando as funções e registrando erros, se houver

try:
    save_query()

except Exception as e:
    logger.error("Erro na execução da função 'save_query()': %s", str(e))

try:
    send_query()

except Exception as e:
    logger.error("Erro na execução da função 'send_query()': %s", str(e))
