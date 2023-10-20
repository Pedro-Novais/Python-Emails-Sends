import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
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
    for i in range(3):
        save_email.append(valor[i][10])
    print(save_email)    

if __name__ == '__main__':
    save()

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def send_test_mail():
    #server = 'smtp.office365.com'
    server = 'smtp-mail.outlook.com'
    port = 587
    sender_email = "phnovaisnew@outlook.com"
    password = "Insano01$"
    for ini in range(3):
        msg = MIMEMultipart('alternative')
        receiver_email = save_email[ini]

        file = read_template("template/ind.txt")
        
        infos = ['','','','','','','','','','','']
        data = 0
        subject = ""
        data_two = 0
        valorB = 0
        valorL = 0
        
        ver_email = 0
        log_email = """<p style="font-family: 'Passions Conflict', cursive;"> Maria Escobar </p>"""
        img_html = """<img src="https://i.postimg.cc/0yNyxsv6/assinatura-JPG.jpg" alt="Assinatura_E-deploy"/>"""
        img_atm = 0

        for i in range(len(infos)):
            infos[i] = valor[ini][i]
            if(i == 2):
                data = format(infos[2], "%d/%m/%Y")
            if(i == 6):
                subject = f'FATURAMENTO E-DEPLOY - {infos[6]}'
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
            if(ver_email == 0):
                for i in range(len(receiver_email)):
                    if(receiver_email[i] == 'g'):
                        if(receiver_email[i+1] == 'm'):
                            if(receiver_email[i+2] == 'a'):
                                if(receiver_email[i+3] == 'i'):
                                    if(receiver_email[i+4] == 'l'):
                                        img_atm = 1
                                        ver_email = 1
                                        log_email = img_html
        blt = infos[1]
        nota = infos[3]
        tmp_html = {
            'ID':infos[0],
            'BLT': blt,
            'DATA_EMISSAO':data,
            'NOTA': nota,
            'RAZAO': infos[4],
            'CNPJ': infos[5],
            'DESCR': infos[6],
            'VALORB': valorB,
            'VALORL': valorL,
            'DATA_VENC': data_two,
            'IMAGEM': log_email
            #'EMAIL': infos[10]
            }
        
        tmp = file.safe_substitute(tmp_html)
        msg.attach(MIMEText(tmp, 'html'))

        title_file = [f'Nota-Fiscal-{nota}.pdf', f'Boleto-{blt}.pdf']
        pasta_blt = f'{blt}.pdf'
        pasta_nota = f'{nota}.pdf'
        pasta_dir_blt = f'Boletos'
        pasta_dir_notas = f'Notas'
        pasta = ""
        pasta_arq = ""
        state = 0
        for num in range(2):
            if(state == 0):
                pasta = pasta_dir_notas
                pasta_arq = pasta_nota
                state = 1
            elif(state == 1):
                pasta = pasta_dir_blt
                pasta_arq = pasta_blt
            pdf = MIMEApplication(open(f'pdf/{pasta}/{pasta_arq}', 'rb').read())
            pdf.add_header('Content-Disposition', 'attachment', filename= title_file[num])
            msg.attach(pdf)

        if(img_atm != 1):
            with open('assinatura.jpg', 'rb') as fp:
                img = MIMEImage(fp.read())
                img.add_header('Content-Disposition', 'attachment', filename="assinatura.jpg")
                msg.attach(img)
      
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
                print(f"E-mail: {receiver_email} enviado,")
        except Exception as e:
            print(f"E-mail: {receiver_email}, falha ao enviar")
            print(e)

if __name__ == '__main__':
    send_test_mail()