import smtplib, ssl
import openpyxl
import time
import os
import sys
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage

def search(title, values):
    for atrb in values:
        print("")
        print(title)
        print(atrb)

def line(qnt_line):
    line = "-"
    for i in range(qnt_line):
        line = line + "-"
    return print(line)

def read_excel():
    try:
        wb = openpyxl.load_workbook('excel/envio-faturamento.xlsx')
    except Exception as e:
        print('')
        print("Erro ao realizar a leitura da planilha, verifique se o nome está correto: 'envio-faturamento.xlsx'")
        print(e)
        sys.exit()
    return wb

wb = read_excel()
aba = wb['TESTE']

valor_title = []
valor = []
save_email = []

#Function to read the excel and save the infos  
def saving_email():
    cont_i = 0
    for linha in aba:
        if(cont_i>0):
            valor.append([])
        for celula in linha:
            posC = celula.column
            posR = celula.row
            valorT = celula.value
            if(posR < 2):
                valor_title.append(valorT)
                cont_i = 1
            if(posR > 1):
                valor[posR-2].append(valorT)

#Loop to save email in a only variable:
    for i in range(len(valor)):
        save_email.append(valor[i][10])

#Creation variables of save emails equals
email_rep = []
valor_rep = []
qnt_rep = []
cont_qnt = 0
cont = 0
before_email = 0
now_email = 0
index_del = []

#Functions to save the emails equals
def saving_email_rep():
    back = 0
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
            if(i + 1 == len(save_email)):
                cont = i - 1 
                email_rep.append(valor[cont][10])
                valor_rep.append(valor[cont + 1])
                index_del.append(cont + 1)

    for i in range(len(email_rep)):  
        if(i == 0):
            cont_qnt = 2
        if(i > 1):
            decre = i - 1
            ulti = i + 1
            if(email_rep[i] == email_rep[decre]):
                cont_qnt = cont_qnt + 1
            if(email_rep[i] != email_rep[decre]):
                qnt_rep.append([cont_qnt])
                cont_qnt = 1
            if(ulti == len(email_rep)):
                qnt_rep.append([cont_qnt]) 

    index_del.reverse() 

    for i in range(len(index_del)):
        del save_email[index_del[i]]
        del valor[index_del[i]]

#Process to execute the functions to read the Excel and save the emails
if __name__ == '__main__':
   saving_email()
   saving_email_rep()

#Variable that in case of any archive empty, it'll save this archives and show to user
arq_falta = []

#Function to verify if exist every archives in excel
def verification_pdf():
    exi = 0
    dir_arq = ['Boletos', 'Notas']
    qnt_analysi = 0
    for indice_i in range(2):
        if(indice_i == 0):
            qnt_analysi = valor
        elif(indice_i == 1):
            qnt_analysi = valor_rep
        for i in range(len(qnt_analysi)):
            blt_not = []

            blt_not.append(qnt_analysi[i][1])
            blt_not.append(qnt_analysi[i][3])

            line(100)
            for ini in range(2):
                dir = f'pdf/{dir_arq[ini]}/{blt_not[ini]}.pdf'
                if os.path.exists(dir):
                    print(f'{dir_arq[ini]}, de número {blt_not[ini]} existe')
                else:
                    print(f'{dir_arq[ini]}, de número {blt_not[ini]} não existe')
                    arq_falta.append([dir_arq[ini], blt_not[ini]])
                    exi = -1 
    return exi

#Function to read and transform the archive in a Template of email
def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

#Executes the function verification_pdf()
if __name__ == '__main__':
   stop = verification_pdf()

#Function that make the all analysis and send the email
def send_email():
    if(stop == 0):
        email = ""
        email_send = ""
        verify_alert = 1
        info_rep = []
        blt_rep = []
        nota_rep = []
        for i_first in range(2):
            #i_first = 1
            if(i_first == 0):
                email = len(qnt_rep) 
                email_send = email_rep[i_first]
            else:
                email = len(save_email)
                email_send = save_email
        
            #server = 'smtp.office365.com'
            server = 'smtp-mail.outlook.com'
            port = 587
            sender_email = "phnovaisnew@outlook.com"
            password = "Insano01$"

            for ini in range(email):
                if(i_first == 0):
                    if(ini > 0 ):
                        if(len(qnt_rep) >= 1):
                                for delete in range(len(email_rep)):
                                    if(delete < num_rep):
                                        del email_rep[0]
                                        del valor_rep[0]
                                del qnt_rep[0]
                                if(len(email_rep) > 0):
                                    email_send = email_rep[0]

                msg = MIMEMultipart('alternative')
                if(i_first == 0):
                    num_rep = qnt_rep[0][0]
                if(i_first == 1):
                    receiver_email = email_send[ini]
                else:
                    receiver_email = email_send

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

                comment = '<!-- -->'
                tmp_add = [
                    comment,comment,comment,comment,comment,comment,comment,comment,comment,comment,
                    comment,comment,comment,comment,comment,comment,comment,comment,comment,comment,
                    comment,comment,comment,comment,comment,comment,comment,comment,comment,comment
                    ]
                info_add = ['','','','','','','','','','']
                if(i_first == 0):
                        for index in range(num_rep):
                            if(index > 0):
                                info_rep.append([])
                                for adding in range(10):
                                    info_add[adding] = valor_rep[index][adding]
                                info_rep[index-1] = info_add
                                #info_rep.append([])
                                for i in range(len(info_add)):
                                    if(i == 2):
                                        data = format(info_add[2], "%d/%m/%Y")
                                    if(i == 6):
                                        subject = f'FATURAMENTO E-DEPLOY'
                                    if(i == 7):
                                        string = str(info_add[7])
                                        change = string.replace('.',',')
                                        if(len(change) < 6):
                                            valorB = f'R${change}0'
                                        else:
                                            valorB = f'R${change}'
                                    if(i == 8):
                                        string = str(info_add[8])
                                        change = string.replace('.',',')
                                        if(len(change) < 6):
                                            valorL = f'R${change}0'
                                        else:
                                            valorL = f'R${change}'
                                    if(i == 9):
                                        data_two = format(info_add[9], "%d/%m/%Y")
                                
                                template_reu = f"""
                                <th style="padding: 0.2rem; background-color:#fff; border-left:solid 1px; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{info_add[0]}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{info_add[1]}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{data}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{info_add[3]}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{info_add[4]}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 1rem; min-width: 10vw">{info_add[5]}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{info_add[6]}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{valorB}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{valorL}</th>
                                <th style="padding: 0.2rem; background-color:#fff; border-right:solid 1px; border-bottom:solid 1px; font-size: 0.9rem">{data_two}</th>
                                """
                                tmp_add[index - 1] = template_reu

                                blt = info_add[1]
                                nota = info_add[3]
                                blt_rep.append(blt)
                                nota_rep.append(nota)

                                title_file = [f'Boleto-{blt}.pdf', f'Nota_Fiscal-{nota}.pdf']
                                pasta_blt = f'{blt}.pdf'
                                pasta_nota = f'{nota}.pdf'
                                pasta_dir_blt = f'Boletos'
                                pasta_dir_notas = f'Notas'
                                pasta = ""
                                pasta_arq = ""

                                for num in range(2):
                                    if(num == 1):
                                        pasta = pasta_dir_notas
                                        pasta_arq = pasta_nota   
                                    elif(num == 0):
                                        pasta = pasta_dir_blt
                                        pasta_arq = pasta_blt


                                    pdf = MIMEApplication(open(f'pdf/{pasta}/{pasta_arq}', 'rb').read())
                                    pdf.add_header('Content-Disposition', 'attachment', filename= title_file[num])
                                    msg.attach(pdf)

                            info_add = ['','','','','','','','','','']
                        for i in range(len(infos)):
                            infos[i] = valor_rep[0][i]
                            if(i == 2):
                                data = format(infos[2], "%d/%m/%Y")
                            if(i == 6):
                                subject = f'FATURAMENTO E-DEPLOY'
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
                if(i_first == 1):
                    for i in range(len(infos)):
                        infos[i] = valor[ini][i]
                        if(i == 2):
                            data = format(infos[2], "%d/%m/%Y")
                        if(i == 6):
                            subject = f'FATURAMENTO E-DEPLOY'
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
                    'INFOS2': tmp_add[0],
                    'INFOS3': tmp_add[1],
                    'INFOS4': tmp_add[2],
                    'INFOS5': tmp_add[3],
                    'INFOS6': tmp_add[4],
                    'INFOS7': tmp_add[5],
                    'INFOS8': tmp_add[6],
                    'INFOS9': tmp_add[7],
                    'INFOS10': tmp_add[8],
                    'INFOS11': tmp_add[9],
                    'INFOS12': tmp_add[10],
                    'INFOS13': tmp_add[11],
                    'INFOS14': tmp_add[12],
                    'INFOS15': tmp_add[13],
                    'INFOS16': tmp_add[14],
                    'INFOS17': tmp_add[15],
                    'INFOS18': tmp_add[16],
                    'INFOS19': tmp_add[17],
                    'INFOS20': tmp_add[18],
                    'INFOS21': tmp_add[19],
                    'INFOS22': tmp_add[20],
                    'INFOS23': tmp_add[21],
                    'INFOS24': tmp_add[22],
                    'INFOS25': tmp_add[23]
                    }
              
                tmp = file.safe_substitute(tmp_html)
                msg.attach(MIMEText(tmp, 'html'))

                title_file = [f'Boleto-{blt}.pdf', f'Nota_Fiscal-{nota}.pdf']
                pasta_blt = f'{blt}.pdf'
                pasta_nota = f'{nota}.pdf'
                pasta_dir_blt = f'Boletos'
                pasta_dir_notas = f'Notas'
                pasta = ""
                pasta_arq = ""      
                
                for num in range(2):
                    if(num == 1):
                        pasta = pasta_dir_notas
                        pasta_arq = pasta_nota
                    elif(num == 0):
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
                
                
                """
                if(verify_alert == 0):
                    print('')
                    con_alerts = input('Deseja ter que fazer a confirmação para enviar o email? Y/N: ')
                    verify_alert = 1
                con_alerts = ""
                """
    
                print('')
                line(100)
                print(f'Email do remetente: {receiver_email}')
                print('')
                ask = input('Deseja verficar as informações a serem enviadas? S/N: ')
                
                if(ask == 'S' or ask == 's' or ask == 'Y' or ask == 'y'):
                    if(i_first == 0):
                        print('')
                        print('Dados a serem enviados como tabela no email: ')
                        line(100)
                        infos.pop(-1)
                        print(infos)
                        print('')
                        print(f'Nº de boleto: {blt}')
                        print(f'Nº de notas fiscais: {nota}')
                        line(100)
                        for i in range(len(info_rep)):
                            #line(100)
                            print(info_rep[i])
                            print('')
                            print(f'Nº de boleto: {blt_rep[i]}')
                            print(f'Nº de notas fiscais: {nota_rep[i]}')
                            line(100)
                        print('')
                    else:
                        print('')
                        print('Dados a serem enviados como tabela no email: ')
                        print('')
                        infos.pop(-1)
                        print(infos)
                        print('')
                        print(f'Nº de boleto: {blt}')
                        print(f'Nº de notas fiscais: {nota}')
                        line(100)
                    print('')
                    input('Digite qualquer tecla para prosseguir: ')
                    print('')
                
                for i in range(len(blt_rep)):
                    del blt_rep[0]
                    del nota_rep[0]
                    del info_rep[0]

                cont = 1

                """
                if(con_alerts == 'Y' or con_alerts == 'y'):
                    print('')
                    confirm = input('Confirmar envio? Y/N: ')
                    if(confirm == 'Y' or confirm == 'y'):
                        cont = 1
                    elif(con_alerts == 'N' or con_alerts == "n"):
                        sys.exit()
                """
                
                if(cont == 1):
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
                            time.sleep(2)
                            
                            print(f"E-mail: {receiver_email} enviado,")
                    except Exception as e:
                        print(f"E-mail: {receiver_email}, falha ao enviar")
                        print(e)
    else:
        print('')
        print(f'Os seguintes arquivos não foram encontrados: {arq_falta}')
        line(100)

#Executes the function send_email()
if __name__ == '__main__':
    send_email()