import smtplib
import email.message
from string import Template


email_contentT = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Table</title>

    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@100&family=Open+Sans&family=Poppins:wght@200;300&display=swap" rel="stylesheet">
    
    <style>
      body table tr:nth-child(1) {
        background-color: red;
      }
      body table th {
        padding: 1rem;
      }
    </style>
</head>
<body>
    <table>
        <tr>
          <th>Id Interno</th>
          <th>N Boleto</th>
          <th>Data de Emissao</th>
          <th>N Nota Fiscal</th>
          <th>Razao Social</th>
          <th>CNPJ</th>
          <th>Descricao do Servico</th>
          <th>Valor Bruto</th>
          <th>Valor Liquido</th>
          <th>Data de Vencimento</th>
          <th>E-mail</th>
        </tr>
        <tr>
            <th>${id}</th>
            <th>${boleto}</th>
            <th>${data}</th>
            <th>${nota}</th>
            <th>${razao}</th>
            <th>${cnpj}</th>
            <th>${descr}</th>
            <th>${valorB}</th>
            <th>${valorL}</th>
            <th>${dataV}</th>
            <th>${email}</th>
        </tr>
      </table>
</body>
</html>
"""

msg = email.message.Message()
msg['Subject'] = 'Tutsplus Newsletter'
msg['From'] = "phnovaisnew@outlook.com"
msg['To'] = 'totempedro941@gmail.com'


password = "Insano01$"
msg.add_header('Content-Type', 'text/html')
msg.set_payload(email_contentT)

s = smtplib.SMTP('smtp.office365.com:587')
s.starttls()
# Login Credentials for sending the mail 
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string())
