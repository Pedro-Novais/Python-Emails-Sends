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
    server = 'smtp.office365.com'
    port = 587
    sender_email = "phnovaisnew@outlook.com"
    password = "Insano01$"
    for ini in range(2):
        msg = MIMEMultipart('alternative')
        receiver_email = save_email[ini]

        file = read_template("template/ind.txt")
        
        infos = ['','','','','','','','','','','']
        data = 0
        subject = ""
        data_two = 0
        valorB = 0
        valorL = 0

        for i in range(len(infos)):
            infos[i] = valor[ini][i]
            if(i == 2):
                data = format(infos[2], "%d/%m/%Y")
            if(i == 6):
                subject = "FATURAMENTO E-DEPLOY - " + infos[6]
            if(i == 7):
                string = str(infos[7])
                change = string.replace('.',',')
                if(len(change) < 6):
                    valorB = f'R${change}0'
                else:
                    valorB = f'R${change}'
            if(i == 8):
                #valorL = infos[8]
                string = str(infos[8])
                change = string.replace('.',',')
                if(len(change) < 6):
                    valorL = f'R${change}0'
                else:
                    valorL = f'R${change}'
            if(i == 9):
                data_two = format(infos[9], "%d/%m/%Y")
        
        tmp_html = {
            'ID':infos[0],
            'BLT':infos[1],
            'DATA_EMISSAO':data,
            'NOTA': infos[3],
            'RAZAO': infos[4],
            'CNPJ': infos[5],
            'DESCR': infos[6],
            'VALORB': valorB,
            'VALORL': valorL,
            'DATA_VENC': data_two,
            #'EMAIL': infos[10]
            }
        
        tmp = file.safe_substitute(tmp_html)
        msg.attach(MIMEText(tmp, 'html'))

        title_file = ["12345.pdf", "boleto.pdf"]
        for num in range(2):
            pdf = MIMEApplication(open(f'pdf/Line_{ini}/QRCode_{num}.pdf', 'rb').read())
            pdf.add_header('Content-Disposition', 'attachment', filename= title_file[num])
            msg.attach(pdf)
      
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = receiver_email

        context = ssl.create_default_context()

        try:
            with smtplib.SMTP(server, port) as smtpObj:
                smtpObj.ehlo()
                smtpObj.starttls(context=context)
                smtpObj.ehlo()
                smtpObj.login(sender_email, password)
                smtpObj.sendmail(sender_email, receiver_email, msg.as_string())
                print(f"E-mail da linha {ini+1} enviado")
        except Exception as e:
            print(f"E-mail da linha {ini+1}, falha ao enviar")
            print(e)

if __name__ == '__main__':
    send_test_mail()