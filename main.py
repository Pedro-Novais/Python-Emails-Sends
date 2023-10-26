import smtplib, ssl
import openpyxl
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from string import Template

wb = openpyxl.load_workbook('excel/teste_faturamento-teste-email-duplicado 1.xlsx')

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
            if(posR < 15):
                valor[posR-2].append(valorT)

for i in range(len(valor)):
    save_email.append(valor[i][10])

email_rep = []
valor_rep = []
qnt_rep = []
cont_qnt = 0
cont = 0
before_email = 0
now_email = 0
back = 0
index_del = []

for i in range(len(save_email)):
    if(i > 0):
        before_email = save_email[i-1]
        now_email = save_email[i]
        if(now_email == before_email):
            cont = i - 1 
            email_rep.append(valor[cont][10])
            valor_rep.append(valor[cont])
            back = back + 1
            index_del.append(cont)
        elif(now_email != before_email):
                if(back > 0):
                    cont = i - 1 
                    email_rep.append(valor[cont][10])
                    valor_rep.append(valor[cont])
                    index_del.append(cont)
                    back = 0

for i in range(len(email_rep)):
    print(i)
    if(i == 0):
        cont_qnt = 2
    if(i > 1):
        decre = i - 1
        ulti = i + 1
        if(email_rep[i] == email_rep[decre]):
            print('meio')
            cont_qnt = cont_qnt + 1
        if(email_rep[i] != email_rep[decre]):
            print('final')
            qnt_rep.append([cont_qnt])
            cont_qnt = 1
        if(ulti == len(email_rep)):
            print('ultima')
            qnt_rep.append([cont_qnt]) 

print(len(email_rep))
#print(valor_rep)
#print('')
print(len(qnt_rep))

index_del.reverse() 
for i in range(len(index_del)):
    print(index_del)
    del save_email[index_del[i]]
    del valor[index_del[i]]

arq_falta = []
def verification_pdf():
    exi = 0
    dir_arq = ['Boletos', 'Notas']
    for i in range(3):
        blt_not = []

        blt_not.append(valor[i][1])
        blt_not.append(valor[i][3])
        for ini in range(2):
            dir = f'pdf/{dir_arq[ini]}/{blt_not[ini]}.pdf'
            if os.path.exists(dir):
                print('')
                print(f'{dir_arq[ini]}, de número {blt_not[ini]} existe')
            else:
                print('')
                print(f'{dir_arq[ini]}, de número {blt_not[ini]} não existe')
                arq_falta.append([dir_arq[ini], blt_not[ini]])
                exi = -1
    return exi

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

if __name__ == '__main__':
   stop = verification_pdf()

print(stop)
print(arq_falta)

#state_email = 0
def send_test_mail():
    if(stop == 0):
        for i_first in range(1):
            email = ""
            email_send = ""
            if(i_first == 0):
                email = len(qnt_rep)
                email_send = email_rep
            else:
                email = len(save_email)
                email_send = save_email
            #server = 'smtp.office365.com'
            server = 'smtp-mail.outlook.com'
            port = 587
            sender_email = "phnovaisnew@outlook.com"
            password = "Insano01$"

            for ini in range(1):

                msg = MIMEMultipart('alternative')
                receiver_email = f't{save_email[ini]}'

                file = read_template("template/ind.txt")
                
                infos = ['','','','','','','','','','','']
                data = 0
                subject = ""
                data_two = 0
                valorB = 0
                valorL = 0
                
                comment = '<!-- -->'
                
                testei = ["","","","","","","","","",""]
                for indice in range(len(testei)):
                    testei[indice] = valor[1][10]
                teste = f"""
                <th style="padding: 0.5rem; background-color:#fff; border-left:solid 1px; border-right:solid 1px; border-bottom:solid 1px;">${testei[0]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[1]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[2]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[3]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[4]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[5]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[6]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[7]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[8]}</th>
                <th style="padding: 0.5rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px;">${testei[9]}</th>
                """

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
                    'IMAGEM': log_email,
                    #'EMAIL': infos[10]
                    'INFOS2': teste,
                    'INFOS3': comment,
                    'INFOS4': comment,
                    'INFOS5': comment,
                    'INFOS6': comment,
                    'INFOS7': comment,
                    'INFOS8': comment
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
        else:
            print(f'Os seguintes arquivos não foram encontrados: {arq_falta}')

if __name__ == '__main__':
    send_test_mail()