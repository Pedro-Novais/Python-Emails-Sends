import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import openpyxl

wb = openpyxl.load_workbook('teste.xlsx')

def search(title, values):
    for atrb in values:
        print("")
        #print(title)
        #print(atrb)
message = "Teste de mensagem automatiza pelo python, bla bla bla bla"
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
        #print(f'Coluna {posC}, linha {posR} contém células: {valor}')
#print(valor_title)
#print(valor)
print(valor[0][2])
search(valor_title, valor)

#smtpObj = smtplib.SMTP('smtp.gmail.com', 465)

#smtpObj.ehlo()
#smtpObj.starttls()
#smtpObj.login("totempedro941@gmail.com", "Insano0$")
def send_test_mail(body):
    sender_email = "phnovaisnew@outlook.com"
    receiver_email = 'totempedro941@gmail.com'
    context = ssl.create_default_context()
    msg = MIMEMultipart()
    msg['Subject'] = 'Cobrança'
    msg['From'] = sender_email
    msg['To'] = receiver_email

    msgText = MIMEText('<b>%s</b>' % (body), 'html')
    msg.attach(msgText)

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
    send_test_mail(message)

"""
emailSender = "phnovaisnew@outlook.com"
myThroaway = "totempedro941@gmail.com"
emailRecipients = [myThroaway]
newEmail = ""From: From Person phnovais7@gmail.com
            To: To Person totempedro941@gmail.com
            Subject: Email Test
            This is the body of the email.
            ""
try:
    smtpObj = smtplib.SMTP('smtp.office365.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login("phnovaisnew@outlook.com", "Insano01$")
    smtpObj.sendmail(emailSender, emailRecipients, newEmail)
    print ("Certo")
except Exception as e:
    print (e)
"""