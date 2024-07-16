import jinja2
import pdfkit
import pandas as pd
from datetime import datetime
import time
import os
import smtplib
from email.message import EmailMessage
import webbrowser

list_email = []
list_nomes = []
list_filename = []
email_file = open('email.txt', 'r')
email_code_file = open('email_code.txt', 'r')
email = email_file.readline()
email_file.close()
senha = email_code_file.readline()
email_code_file.close
def createPDF(contentPDF):
    template_loader = jinja2.FileSystemLoader('.\\assets\\')
    template_env = jinja2.Environment(loader=template_loader)

    html_template = 'template.html'
    template = template_env.get_template(html_template)
    output_text = template.render(context)
    name_pdf = f'pdf_{nome_loca}.pdf'
    list_filename.append(name_pdf)

    config = pdfkit.configuration(wkhtmltopdf='.\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')
    output_pdf = f'.\\pdf\\pdf_{nome_loca}.pdf'
    pdfkit.from_string(output_text, output_pdf, configuration=config)

date_month_year = datetime.today().strftime("%d %b, %Y")
nome_loca = ""
end_loca = ""
reajuste = 0
salario_antigo = ''
salario_novo = ''
data_inicio = ''

context = {
    'date_month_year': date_month_year,
    'nome_loca': nome_loca,
    'end_loca': end_loca,
    'reajuste': reajuste,
    'salario_antigo': salario_antigo,
    'salario_novo': salario_novo,
    'data_inicio': data_inicio
}

datasheet = pd.read_excel(".\\database\\datasheet.xlsx", header=0)
while True:
    os.system('cls')
    print(f"""
    
   #                        #     #                 
  # #   #    # #####  ####  ##   ##   ##   # #      
 #   #  #    #   #   #    # # # # #  #  #  # #      
#     # #    #   #   #    # #  #  # #    # # #      
####### #    #   #   #    # #     # ###### # #      
#     # #    #   #   #    # #     # #    # # #      
#     #  ####    #    ####  #     # #    # # ###### 
          
logado como: {email}
""")
    
    try:
        opt = int(input("deseja qual função?\n[0]Sair\n[1]Enviar para todos\n[2]Enviar para um especifico\n[3] printar Banco Excel (Recomendado so para DEBUG)\n[4] Enviar emails para todos\n[5] Configurar programa\n[6] Tutorial de como conseguir codigo email google\n>>> "))
        if opt == 1:
            reajuste = float(input("digite o valor do reajuste sem o % (ex: 10): "))
            data_inicio = str(input("digite a data de inicio: "))
            for i, row in datasheet.iterrows():
                try:
                    nome_loca = row['Nome']
                    info = (datasheet[datasheet['Nome'] == nome_loca])
                    print(f"criando arquivo para {nome_loca}")
                    list_email.append(info['Email'].values[0])
                    list_nomes.append(nome_loca)
                    end_loca = row['Endereço']
                    salario_antigo = row['aluguel']
                    salario_novo = ((salario_antigo/100)*reajuste) + salario_antigo
                    salario_novo_formatado = "{:.2f}".format(salario_novo)  
                    salario_antigo_formatado = "{:.2f}".format(salario_antigo)
                    datasheet.loc[datasheet['Endereço'] == end_loca, 'aluguel'] = salario_novo
                    datasheet.to_excel('.\\database\\datasheet.xlsx', index=False)
                    context = {
                        'date_month_year': date_month_year,
                        'nome_loca': nome_loca,
                        'end_loca': end_loca,
                        'reajuste': reajuste,
                        'salario_antigo': salario_antigo_formatado,
                        'salario_novo': salario_novo_formatado,
                        'data_inicio': data_inicio
                    }
                    createPDF(context)
                except Exception as e:
                    print(e)
                    input()
        elif opt == 2:
        #------------------INPUT ONE BY ONE-----------------------
            inputNome = str(input("Digite o nome para quem o email sera escrito: "))


            info = (datasheet[datasheet['Nome'] == inputNome])
            if not info.empty:
                print(info['Nome'].values[0])
                print(info['Endereço'].values[0])
                print(info['aluguel'].values[0])
                confirm = int(input("As informações estão corretas?: [0]SIM [1]NÃO :"))
                if confirm == 0:
                    nome_loca = info['Nome'].values[0]
                    end_loca = info['Endereço'].values[0]
                    salario_antigo = info['aluguel'].values[0]

                    reajuste = float(input("digite a porcentagem de mudança sem  % (ex: 10): "))
                    salario_novo = ((salario_antigo/100)*reajuste) + salario_antigo
                    salario_antigo_formatado = "{:.2f}".format(salario_antigo)
                    salario_novo_formatado = "{:.2f}".format(salario_novo)
                    datasheet.loc[datasheet['Endereço'] == end_loca, 'aluguel'] = salario_novo
                    datasheet.to_excel('.\\database\\datasheet.xlsx', index=False)
                    data_inicio = str(input("digite a data onde a mudança começará: "))
                    context = {
                        'date_month_year': date_month_year,
                        'nome_loca': nome_loca,
                        'end_loca': end_loca,
                        'reajuste': reajuste,
                        'salario_antigo': salario_antigo_formatado,
                        'salario_novo': salario_novo_formatado,
                        'data_inicio': data_inicio
                    }
                    createPDF(context)
            else:
                print("Nome não encontrado!")
                print("Por favor confira o banco de dados para ver se o nome foi escrito de maneira correta.")
        elif opt == 3:
            print(datasheet)
            input("")
        elif opt == 4:
            server_type = 'smtp-mail.outlook.com'
            port = 587
            for conta in range(len(list_nomes)):
                try:
                    print(f"nome {list_nomes[conta]} EMAIL: {list_email[conta]} FILE: {list_filename[conta]}")
                    msg = EmailMessage()
                    msg['From'] = email
                    msg.set_content("Aqui estão as novas atualizações em relação a mudança do aluguel de acordo com as normas IGP-M.")
                    msg['To'] = list_email[conta]
                    msg['Subject'] = "Atualização do valor de custo mensal da propriedade."
                    fileToSend = list_filename[conta]
                    emailToSend = list_email[conta]
                    print(emailToSend)
                    contFile = open(f'.\\pdf\\{fileToSend}', 'rb')
                    cont = contFile.read()
                    msg.add_attachment(cont, maintype='application', subtype='pdf', filename= fileToSend)
                    server = smtplib.SMTP(server_type, port)
                    server.starttls()
                    server.login(email, senha)
                    server.send_message(msg)
                    msg.clear_content()
                except Exception as e:
                    print(f"erro encontrado favor contate suporte caso necessario \n {e}")
                    input("")

        elif opt == 5:
            email = str(input("Digite o seu email que sera usado para enviar os documentos para os moradores: "))
            email_code = str(input("Digite a senha do email: "))
            email_file = open('email.txt', 'w')
            email_file.write(email)
            email_code_file = open('email_code.txt', 'w')
            email_code_file.write(email_code)
            email_file.close()
            email_code_file.close()
        elif opt == 0:
            break
    except Exception as e:
        print("Erro encontrado favor consulte o suporte caso necessario\n {e}")
