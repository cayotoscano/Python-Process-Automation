# importações
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pathlib
import os

# importar dataframes
diretorio_atual = os.getcwd()
vendas = pd.read_excel(diretorio_atual + r"\Bases de Dados\Vendas.xlsx")
lojas = pd.read_csv(diretorio_atual + r"\Bases de Dados\Lojas.csv", encoding="latin1", sep=";")
emails = pd.read_excel(diretorio_atual + r"\Bases de Dados\Emails.xlsx")

# unificar os dataframes
vendas = vendas.merge(lojas, on="ID Loja")

# criar um dataframe para cada loja
dicionario_lojas = {}
for loja in lojas["Loja"]:
    dicionario_lojas[loja] = vendas.loc[vendas["Loja"]==loja, :]

# definir dia de consulta
dia_indicador = vendas["Data"].max()

# salvar backups de cada loja em uma pasta no formato "mês_dia_NomeLoja.xlsx"
caminho_backup = pathlib.Path(diretorio_atual + r"\Backup Arquivos Lojas")
for loja in lojas["Loja"]:
    if not pathlib.Path(caminho_backup, loja).exists():
        pathlib.Path(caminho_backup, loja).mkdir()

    nome_arquivo = f"{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx"
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)

# metas gerais das lojas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1_650_000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

# calcular indicadores de cada uma das lojas e enviar E-mail para os gerentes
for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja["Data"]==dia_indicador, :]

    # faturamento
        # dia
    faturamento_dia = vendas_loja_dia["Valor Final"].sum()
        # ano
    faturamento_ano = vendas_loja["Valor Final"].sum()

    # diversidade de produtos
        # dia
    quantiade_produtos_dia = len(vendas_loja_dia["Produto"].unique())
        # ano
    quantiade_produtos_ano = len(vendas_loja["Produto"].unique())

    # ticket médio
        # dia
    tabela_ticket_medio_dia = vendas_loja_dia.groupby("Código Venda")["Valor Final"].sum().reset_index()
    ticket_medio_dia = tabela_ticket_medio_dia["Valor Final"].mean()
        # ano
    tabela_ticket_medio = vendas_loja.groupby("Código Venda")["Valor Final"].sum().reset_index()
    ticket_medio_ano = tabela_ticket_medio["Valor Final"].mean()

    # enviar E-mail
    sender_email = 'cayocraft@gmail.com'
    sender_password = 'senha' 
    nome = emails.loc[emails["Loja"]==loja, 'Gerente'].values[0]
    recipient_email = emails.loc[emails["Loja"]==loja, 'E-mail'].values[0]
    attachment_path = caminho_backup / loja / f"{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx"  # caminho do arquivo
    attachment_path = str(attachment_path)

    # definir cor de conclusão das metas (verde para metas batidas e vermelho para o oposto)
    if faturamento_dia >= meta_faturamento_dia:
        cor_faturamento_dia = "green"
    else:
        cor_faturamento_dia = "red"
    if faturamento_ano >= meta_faturamento_ano:
        cor_faturamento_ano = "green"
    else:
        cor_faturamento_ano = "red"

    if quantiade_produtos_dia >= meta_qtdeprodutos_dia:
        cor_quantidade_dia = "green"
    else:
        cor_quantidade_dia = "red"
    if quantiade_produtos_ano >= meta_qtdeprodutos_ano:
        cor_quantidade_ano = "green"
    else:
        cor_quantidade_ano = "red"

    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = "green"
    else:
        cor_ticket_dia = "red"
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = "green"
    else:
        cor_ticket_ano = "red"

    # corpo do E-mail em HTML
    subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    body = f"""
    <p>Bom dia, {nome}</p>
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da loja <strong>{loja}</strong> foi:</p>

    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
    </tr>
    <tr>
        <td style='text-align: center'>Faturamento</td>
        <td style='text-align: center'>R${faturamento_dia:.2f}</td>
        <td style='text-align: center'>R${meta_faturamento_dia:.2f}</td>
        <td style='text-align: center'><font color='{cor_faturamento_dia}'>◙</font></th>
    </tr>
    <tr>
        <td style='text-align: center'>Diversidade de Produtos</td>
        <td style='text-align: center'>{quantiade_produtos_dia}</td>
        <td style='text-align: center'>{meta_qtdeprodutos_dia}</td>
        <td style='text-align: center'><font color='{cor_quantidade_dia}'>◙</font></th>
    </tr>
    <tr>
        <td style='text-align: center'>Ticket Médio</td>
        <td style='text-align: center'>R${ticket_medio_dia:.2f}</td>
        <td style='text-align: center'>R${meta_ticketmedio_dia:.2f}</td>
        <td style='text-align: center'><font color='{cor_ticket_dia}'>◙</font></th>
    </tr>
    </table>
    <br>
    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
    </tr>
    <tr>
        <td style='text-align: center'>Faturamento</td>
        <td style='text-align: center'>R${faturamento_ano:.2f}</td>
        <td style='text-align: center'>R${meta_faturamento_ano:.2f}</td>
        <td style='text-align: center'><font color='{cor_faturamento_ano}'>◙</font></th>
    </tr>
    <tr>
        <td style='text-align: center'>Diversidade de Produtos</td>
        <td style='text-align: center'>{quantiade_produtos_ano}</td>
        <td style='text-align: center'>{meta_qtdeprodutos_ano}</td>
        <td style='text-align: center'><font color='{cor_quantidade_ano}'>◙</font></th>
    </tr>
    <tr>
        <td style='text-align: center'>Ticket Médio</td>
        <td style='text-align: center'>R${ticket_medio_ano:.2f}</td>
        <td style='text-align: center'>R${meta_ticketmedio_ano:.2f}</td>
        <td style='text-align: center'><font color='{cor_ticket_ano}'>◙</font></th>
    </tr>
    </table>


    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Cayo</p>
    """

    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_email
    message['Subject'] = subject

    html_part = MIMEText(body, 'html')
    message.attach(html_part)

    # anexos
    try:
        with open(attachment_path, 'rb') as attachment_file:
            attachment_part = MIMEBase('application', 'octet-stream')
            attachment_part.set_payload(attachment_file.read())
            encoders.encode_base64(attachment_part)
            attachment_part.add_header(
                'Content-Disposition',
                f'attachment; filename="{dia_indicador.day}_{dia_indicador.month}_d.xlsx"'
            )
            message.attach(attachment_part)
    except:
        pass

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [recipient_email], message.as_string())

# criar ranking para diretoria e salvar backup
faturamento_lojas_ano = vendas.groupby("Loja")[["Valor Final"]].sum()
faturamento_lojas_ano = faturamento_lojas_ano.sort_values(by="Valor Final", ascending=False)

nome_arquivo = f"{dia_indicador.day}_{dia_indicador.month}_Ranking Anual.xlsx"
faturamento_lojas_ano.to_excel(r'C:\Users\cayo_\Desktop\PROGRAMAÇÃO\PROJETOS\PROJETO 1\Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas["Data"]==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby("Loja")[["Valor Final"]].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by="Valor Final", ascending=False)

nome_arquivo = f"{dia_indicador.day}_{dia_indicador.month}_Ranking Diário.xlsx"
faturamento_lojas_dia.to_excel(r'C:\Users\cayo_\Desktop\PROGRAMAÇÃO\PROJETOS\PROJETO 1\Backup Arquivos Lojas\{}'.format(nome_arquivo))

# enviar E-mail para diretoria
sender_email = 'cayocraft@gmail.com'
sender_password = 'senha' 
nome = emails.loc[emails["Loja"]==loja, 'Gerente'].values[0]
recipient_email = emails.loc[emails["Loja"]=="Diretoria", 'E-mail'].values[0]

subject = f'Ranking dia {dia_indicador.day}/{dia_indicador.month}'
body = f"""
<p>Prezados, bom dia</p>
<br>

<p>Melhor loja do dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0,0]:.2f}</p>
<p>Pior loja do dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Fatuaramento R${faturamento_lojas_dia.iloc[-1,0]:.2f}</p>
<p>Melhor loja do ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0,0]:.2f}</p>
<p>Pior loja do ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Fatuaramento R${faturamento_lojas_ano.iloc[-1,0]:.2f}</p>

<br>
<p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
<p>Qualquer dúvida estou à disposição.</p>
<p>Att., Cayo</p>
"""

message = MIMEMultipart()
message['From'] = sender_email
message['To'] = recipient_email
message['Subject'] = subject

html_part = MIMEText(body, 'html')
message.attach(html_part)

attachment_path = caminho_backup / f"{dia_indicador.day}_{dia_indicador.month}_Ranking Anual.xlsx"  # caminho do arquivo
attachment_path = str(attachment_path)

with open(attachment_path, 'rb') as attachment_file:
    attachment_part = MIMEBase('application', 'octet-stream')
    attachment_part.set_payload(attachment_file.read())
    encoders.encode_base64(attachment_part)
    attachment_part.add_header(
        'Content-Disposition',
        f'attachment; filename="{dia_indicador.day}_{dia_indicador.month}_a.xlsx"'
    )
    message.attach(attachment_part)
    print("anexado")


attachment_path = caminho_backup / f"{dia_indicador.day}_{dia_indicador.month}_Ranking Diário.xlsx"  # caminho do arquivo
attachment_path = str(attachment_path)

try:
    with open(attachment_path, 'rb') as attachment_file:
        attachment_part = MIMEBase('application', 'octet-stream')
        attachment_part.set_payload(attachment_file.read())
        encoders.encode_base64(attachment_part)
        attachment_part.add_header(
            'Content-Disposition',
            f'attachment; filename="{dia_indicador.day}_{dia_indicador.month}_d.xlsx"'
        )
        message.attach(attachment_part)
except:
    pass

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, [recipient_email], message.as_string())