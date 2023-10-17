import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from string import Template
import openpyxl
import email.message

wb = openpyxl.load_workbook('teste.xlsx')

def search(title, values):
    for atrb in values:
        print("")
        print(title)
        print(atrb)

aba = wb['Planilha1']
valor_title = []
valor = []
i = 0
for linha in aba:
    if(i>0):
        valor.append([])
    for celula in linha:
        posC = celula.column
        posR = celula.row
        valorT = celula.value
        if(posR < 2):
            valor_title.append(valorT)
            i = 1
        if(posR > 1):
            valor[posR-2].append(valorT)   

if __name__ == '__main__':
    search(valor_title, valor)

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def send_test_mail():
    sender_email = "phnovaisnew@outlook.com"
    receiver_email = 'totempedro941@gmail.com'
    
    context = ssl.create_default_context()
    
    msg = MIMEMultipart('alternative')
    #msg = email.message.Message()
    msg['Subject'] = 'Cobrança'
    msg['From'] = sender_email
    msg['To'] = receiver_email
    #msg.add_header('Content-Type', 'text/html')
    #msg.set_payload(msgt)  

    file = read_template("ind.txt")

    message = file.substitute(id='pedroooooooooooooooooooooooooo')
    """message = file.substitute(boleto='BoletoTeste')
    message = file.substitute(data='Datateste de tamanho')
    message = file.substitute(nota='notateste')
    message = file.substitute(razao='teste')
    message = file.substitute(cnpj='cnteste')
    message = file.substitute(descr='de')
    message = file.substitute(valorB='teste')
    message = file.substitute(valorL='teste')
    message = file.substitute(dataV='t')
    message = file.substitute(email='teste')"""
    print(message)
    
    msg.attach(MIMEText(message, 'html'))
   
    
    pdf = MIMEApplication(open('QRCode.pdf', 'rb').read())
    pdf.add_header('Content-Disposition', 'attachment', filename= "QRCode.pdf")
    msg.attach(pdf)

    try:
        with smtplib.SMTP('smtp.office365.com', 587) as smtpObj:
            smtpObj.ehlo()
            smtpObj.starttls(context=context)
            smtpObj.login("phnovaisnew@outlook.com", "Insano01$")
            smtpObj.sendmail(sender_email, receiver_email, msg.as_string())
            print('Enviado')
    except Exception as e:
        print(e)

if __name__ == '__main__':
    send_test_mail()