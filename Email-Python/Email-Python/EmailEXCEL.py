import tkinter as tk
from tkinter import filedialog
import threading
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

caminho_arquivo = ""
contagem_emails_supervisor = 0
contagem_emails_coordenador = 0

def selecionar_arquivo_e_enviar():
    global caminho_arquivo
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho_arquivo:
        email_thread = threading.Thread(target=enviar_emails, args=(caminho_arquivo,))
        email_thread.start()
    else:
        print("Nenhum arquivo selecionado.")

def enviar_emails(caminho_arquivo):
    
    if caminho_arquivo:
        global contagem_emails_supervisor
        global contagem_emails_coordenador

        contagem_emails_supervisor = 0
        contagem_emails_coordenador = 0

        EnviarEmailSupervisores(caminho_arquivo)
        EnviarEmailCoordenadores(caminho_arquivo)
        EnviarEmailGerenteKevin(caminho_arquivo, contagem_emails_supervisor, contagem_emails_coordenador)

        janela.destroy()

def EnviarEmailSupervisores(caminho_arquivo):
    if caminho_arquivo:
        global contagem_emails_supervisor
        df = pd.read_excel(caminho_arquivo)
        for i, supervisor in df.iterrows():
            msg = MIMEMultipart()
            msg['Subject'] = "Pendências de Aplicação de Feedback! iFood CX - Alpha E Loyalty"
            messageemail = f"""\
                <html>
                <body>
                <p><h4>Olá <font color="red"><strong>{supervisor['Superior']},</font></strong> tudo bem? Espero que sim!</h4></p>
                <p><h4>Você tem um feedback pendente de aplicação na ferramenta 2CLIX. Peço que verifique e garanta o lançamento antes da expiração do prazo do feedback. Qualquer dúvida, estou a disposição.</h4></p>
                <br>
                <p><h5>Código da monitoria: {supervisor['Código da avaliação']}</h5></p>
                <p><h5>Data do prazo do feedback: {supervisor['Data prazo de feedback']}</h5></p>
                <p><h5>Status do feedback: {supervisor['Status do feedback']}</h5></p>
                <p><h5>Operadora(a) monitorado(a): {supervisor['Avaliado']}</h5></p>
                <p><h5>Coordenador(a) de operações: {supervisor['Coordenador']}</h5></p>
                <br>
                <p>Atenciosamente</p>
                <p>Gerencia Kevin Lima, Coordenação Alpha e Loyalty</p>
                <p>Unidade Arapiraca</p>
                <img src="https://i.imgur.com/74arG9m.jpg" />
                </body>
                </html>
                """
            msg['From'] = 'alphaeloyalty@gmail.com'
            msg['To'] = supervisor['Email Supervisor']
            msg.attach(MIMEText(messageemail, 'html'))
            context = smtplib.SMTP('smtp.gmail.com', 587)
            context.starttls()
            context.login('alphaeloyalty@gmail.com', 'zhdgzkgbcnumnopm')
            try:
                context.sendmail(msg['From'], msg['To'], msg.as_string())
                contagem_emails_supervisor += 1
                print(f"Email enviado para {msg['To']} com sucesso.")
            except Exception as e:
                print(f"Erro ao enviar email para {msg['To']}: {str(e)}")

def EnviarEmailCoordenadores(caminho_arquivo):
    if caminho_arquivo:
        global contagem_emails_coordenador
        df = pd.read_excel(caminho_arquivo)
        for i, coordenador in df.iterrows():
            msg = MIMEMultipart()
            msg['Subject'] = "Pendências de Aplicação de Feedback! iFood CX - Alpha E Loyalty"
            messageemail = f"""\
                <html>
                <body>
                <p><h4>Olá <font color="red"><strong>{coordenador['Coordenador']},</font></strong> tudo bem? Espero que sim!</h4></p>
                <p><h4>Seu supervisor <font color="red">{coordenador['Superior']}</font> tem uma monitoria pendente de aplicação de feedback no 2CLIX. Peço que garante a aplicação do feedback na ferramenta antes do prazo de expiração do feedback</h4></p>
                <br>
                <p><h5>Código da monitoria: {coordenador['Código da avaliação']}</h5></p>
                <p><h5>Data do prazo do feedback: {coordenador['Data prazo de feedback']}</h5></p>
                <p><h5>Status do feedback: {coordenador['Status do feedback']}</h5></p>
                <p><h5>Operadora(a) monitorado(a): {coordenador['Avaliado']}</h5></p>
                <p><h5>Supervisor(a): {coordenador['Superior']}</h5></p>
                <br>
                <p>Atenciosamente</p>
                <p>Gerencia Kevin Lima, Coordenação Alpha e Loyalty</p>
                <p>Unidade Arapiraca</p>
                <img src="https://i.imgur.com/74arG9m.jpg" />
                </body>
                </html>
                """
            msg['From'] = 'alphaeloyalty@gmail.com'
            msg['To'] = coordenador['Email Coordenador']
            msg.attach(MIMEText(messageemail, 'html'))
            context = smtplib.SMTP('smtp.gmail.com', 587)
            context.starttls()
            context.login('alphaeloyalty@gmail.com', 'zhdgzkgbcnumnopm')
            try:
                context.sendmail(msg['From'], msg['To'], msg.as_string())
                contagem_emails_coordenador += 1
                print(f"Email enviado para {msg['To']} com sucesso.")
            except Exception as e:
                print(f"Erro ao enviar email para {msg['To']}: {str(e)}")

def EnviarEmailGerenteKevin(caminho_arquivo, contagem_supervisores, contagens_coordenadores):
    if caminho_arquivo:
        df = pd.read_excel(caminho_arquivo)
        gerente_email = 'jose.raimundo@aec.com.br'
        msg = MIMEMultipart()
        msg['Subject'] = "Relatório de Monitorias pendentes de aplicação de feedback"

        contagens_coordenadores = {}
        contagem_supervisores = {}

        for i, coordenador in df.iterrows():
            email_coordenador = coordenador['Email Coordenador']
            if email_coordenador in contagens_coordenadores:
                contagens_coordenadores[email_coordenador] += 1
            else:
                contagens_coordenadores[email_coordenador] = 1
        
        for i, supervisor in df.iterrows():
            email_supervisor = supervisor['Email Supervisor']
            if email_supervisor in contagem_supervisores:
                contagem_supervisores[email_supervisor] += 1
            else:
                contagem_supervisores[email_supervisor] = 1
         
        email_contagem_coordenador = "<font size='3px' color='red'><strong>Contagem de monitorias pendente de aplicação de feedback por coordenador(a):</strong></font>\n"
        for email_coordenador, contagem in contagens_coordenadores.items():
            email_contagem_coordenador += f"{email_coordenador}: {contagem}\n"
        
        email_contagem_supervisores = "<font size='3px' color='red'><strong>Contagem de monitorias pendente de aplicação de feedback por supervisor(a):</strong></font>\n"
        for email_supervisor, contagem in contagem_supervisores.items():
            email_contagem_supervisores += f"{email_supervisor}: {contagem}\n"
        
        messageemail = f"""\
        <html>
        <body>
        <p><h4>Olá, Rai.</h4></p>
        <p><h4>Segue relatório de quantidade de monitorias por supervisor e coordenador da gerencia de Alpha e Loyalty.</h4></p>
        <br>
        <pre>{email_contagem_coordenador}</pre>
        <br>
        <pre>{email_contagem_supervisores}</pre>
        <br>
        <p>Atenciosamente</p>
        <p>Gerencia Rai, Coordenação Alpha e Loyalty</p>
        <img src="https://i.imgur.com/74arG9m.jpg" />
        </body>
        </html>
        """

        msg['From'] = 'alphaeloyalty@gmail.com'
        msg['To'] = gerente_email
        msg.attach(MIMEText(messageemail, 'html'))
        context = smtplib.SMTP('smtp.gmail.com', 587)
        context.starttls()
        context.login('alphaeloyalty@gmail.com', 'zhdgzkgbcnumnopm')
        try:
            context.sendmail(msg['From'], [msg['To']], msg.as_string())
            print(f"Email enviado para Kevin Lima com sucesso.")
        except Exception as e:
            print(f"Erro ao enviar email para Rai: {str(e)}")

janela = tk.Tk()
janela.geometry("300x70")
janela.title("Enviar Email")

botao_selecionar_enviar = tk.Button(janela, text="Selecionar export do 2CLIX", command=selecionar_arquivo_e_enviar, bg="black", fg="white", font="-weight bold -size 15")
botao_selecionar_enviar.pack(padx=15, pady=20)

janela.mainloop()