import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from string import Template
import openpyxl

wb = openpyxl.load_workbook('excel/teste_faturamento.xlsx')

def search(title, values):
    for atrb in values:
        print("")
        print(title)
        print(atrb)

aba = wb['TESTE']
valor_title = []
valor = []
valor_teste = []
save_email = []
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
            if(posR < 9):
                valor[posR-2].append(valorT)

def save():
    for i in range(2):
        save_email.append(valor[i][10])
    print(save_email)    

if __name__ == '__main__':
    save()

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def send_test_mail():
    for ini in range(2):
        sender_email = "phnovaisnew@outlook.com"
        password = "Insano01$"
        receiver_email = save_email[ini]

        subject = ""
        
        context = ssl.create_default_context()

        file = read_template("template/ind.txt")
        
        infos = ['','','','','','','','','','','']
        data = 0
        data_two = 0
        for i in range(len(infos)):
            infos[i] = valor[ini][i]
            if(i == 2):
                data = format(infos[2], "%d/%m/%Y")
            if(i == 6):
                subject = "FATURAMENTO E-DEPLOY - " + infos[6]
            if(i == 9):
                data_two = format(infos[9], "%d/%m/%Y")
        tmp_teste = {
            'ID':infos[0],
            'BLT':infos[1],
            'DATA_EMISSAO':data,
            'NOTA': infos[3],
            'RAZAO': infos[4],
            'CNPJ': infos[5],
            'DESCR': infos[6],
            'VALORB': infos[7],
            'VALORL': infos[8],
            'DATA_VENC': data_two,
            #'EMAIL': infos[10]
            }
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = receiver_email

        tmp = file.safe_substitute(tmp_teste)
        msg.attach(MIMEText(tmp, 'html'))

        pdf = MIMEApplication(open('pdf/QRCode.pdf', 'rb').read())
        pdf.add_header('Content-Disposition', 'attachment', filename= "QRCode.pdf")
        msg.attach(pdf)

        try:
            with smtplib.SMTP('smtp.office365.com', 587) as smtpObj:
                smtpObj.ehlo()
                smtpObj.starttls(context=context)
                smtpObj.login(sender_email, password)
                smtpObj.sendmail(sender_email, receiver_email, msg.as_string())
                print(i,"º email - Enviado")
        except Exception as e:
            print(i,"º email - Erro ao enviar")
            print(e)

if __name__ == '__main__':
    send_test_mail()