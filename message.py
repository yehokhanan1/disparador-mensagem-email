import os
import csv
import sys
import smtplib
import tkinter as tk
from tkinter import *
from csv import reader
from tkinter import filedialog
from email.message import EmailMessage


windonws = Tk()
windonws.title("Message Generator")
windonws.geometry("300x300")
windonws.configure(background="#1C1C1C")
windonws.wm_resizable(width=False, height=False)

logoMG = Label(windonws, text="MG", background='#1C1C1C', foreground='white')
logoMG.configure(font=("Calibri Light", 50, "normal"))
logoMG.place(x=25,y=1, width=250)

linhaBlack = Label(windonws, text="", background='white', foreground='white')
linhaBlack.place(x=65,y=75, width=165,height=2)

MG = Label(windonws, text="Message Generator", background='#1C1C1C', foreground='white')
MG.configure(font=("Calibri Light", 15, "normal"))
MG.place(x=25,y=80, width=250,height=23)

useText = Label(windonws, text="Usuário", background='#1C1C1C', foreground='white', anchor=W)
useText.configure(font=("default", 12, "bold"))
useText.place(x=8,y=127)

passText = Label(windonws, text="Senha", background='#1C1C1C', foreground='white', anchor=W)
passText.configure(font=("default", 12, "bold"))
passText.place(x=8,y=186)

user1 = Entry(windonws)
user1.configure(font=("default", 10, "normal"))
user1.place(x=10,y=150,width=280,height=20)

passwd = Entry(windonws, show="*")
passwd.configure(font=("default", 12, "bold"))
passwd.place(x=10,y=210,width=280,height=20)

user = '123456'
passwordUser = '123456'

try:
    # create paste
    os.makedirs("C:/Users/Emails")
except:
    pass

def userLogin():
    if user == user1.get() and passwordUser == passwd.get():
        sendEmail()
        windonws.destroy()    
    else:
        fail = Label(windonws, text="Usuario ou Senha incorreto", background='red', foreground='white')
        fail.configure(font=("Arial", 10, "normal"))
        fail.place(x=10,y=270, width=280, height=20)

def menubar(janela):
    menubar = Menu(janela)
    filemenu = Menu(menubar, tearoff=0)
    filemenu.add_command(label="Enviar E-mails", command=sendEmail)
    filemenu.add_command(label="Cadastrar E-mails", command=cadastrar)
    filemenu.add_command(label="Criar Banco de Dados", command=criarBanco)

    filemenu.add_separator()

    filemenu.add_command(label="Sair", command=janela.quit)
    menubar.add_cascade(label="Menu Options", menu=filemenu)
    janela.config(menu=menubar)

def cadastrar():
    windonws1 = Tk()
    windonws1.title('Cadastrar Emails')
    windonws1.geometry("330x150")
    windonws1.wm_resizable(width=False, height=False)
    windonws1.configure(background="#1C1C1C")

    menubar(windonws1)

    Email = Label(windonws1, text="Nome do E-mail", background="#1C1C1C",foreground="white", anchor=W)
    Email.configure(font=("default", 10, "bold"))
    Email.place(x=8,y=9)

    EmailAddress1 = Entry(windonws1)
    EmailAddress1.configure(font=("default", 10, "normal"))
    EmailAddress1.place(x=10,y=30,width=300,height=20)
    try:
        pasta = 'C:/Users/Emails'
        for diretorio, subpastas, arquivos in os.walk(pasta):
            arquivo = arquivos

        options_list = arquivo
        value_inside = tk.StringVar(windonws1)
        value_inside.set("Selecione o Banco de Dados")
        question_menu = tk.OptionMenu(windonws1, value_inside, *options_list)
        question_menu.place(x=9,y=55)
        
        def cadastrarEmail():
            try:
                with open(f'{pasta}/{value_inside.get()}', 'r') as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')

                    for linha in csv_reader:
                        GetEmail = EmailAddress1.get()
                    try:
                        IndexEmail = linha.index(GetEmail)

                        messageSucess = f"E-mail já cadastrado em {value_inside.get()}".replace(".csv", "")
                        Cadastrado = Label(windonws1, text=messageSucess, background="red",foreground="white")
                        Cadastrado.configure(font=("default", 10, "normal"))
                        Cadastrado.place(x=10,y=115, width=300, height=25)
                    except:
                        with open(f'{pasta}/{value_inside.get()}', 'r') as file:
                            header = csv.reader(file, delimiter=',')
                            for linha in header:
                                strLinha = str(linha)
                                linhaClear = strLinha.replace("]", "").replace("'", "").replace("[", "")+EmailAddress1.get()+';'.rstrip('\n')
                                listLinha = [linhaClear]
                            with open(f'{pasta}/{value_inside.get()}', 'w', newline='') as file:
                                writer = csv.writer(file)

                                writer.writerow(listLinha)

                        messageFail = f"E-mail cadastrado em {value_inside.get()}".replace(".csv", "")
                        NoCadastrado = Label(windonws1, text=messageFail, background="green",foreground="white")
                        NoCadastrado.configure(font=("default", 10, "normal"))
                        NoCadastrado.place(x=10,y=115, width=300, height=25)
            except:
                Selecionar = Label(windonws1, text="Selecione um banco de dados", background="red",foreground="white")
                Selecionar.configure(font=("default", 10, "normal"))
                Selecionar.place(x=10,y=115, width=300, height=25)

        def removeEmail():
            with open(f'{pasta}/{value_inside.get()}', 'r') as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')

                for linha in csv_reader:
                    GetEmail = EmailAddress1.get()
                    try:
                        while True:
                            linha.remove(GetEmail)
                        
                    except:
                        with open(f'{pasta}/{value_inside.get()}', 'w', newline='') as file:
                            writer = csv.writer(file, delimiter=';')

                            writer.writerow(linha)

                messageRemove = f"E-mail removido em {value_inside.get()}".replace(".csv", "")
                remove = Label(windonws1, text=messageRemove, background="green",foreground="white")
                remove.configure(font=("default", 10, "normal"))
                remove.place(x=10,y=115, width=300, height=25)        

        CadastrarButton = Button(windonws1, text="Cadastrar", background="#A4A4A4",foreground="black", command=cadastrarEmail)
        CadastrarButton.configure(font=("default", 10, "bold"))
        CadastrarButton.place(x=10,y=90, width=75, height=20)

        deletButton = Button(windonws1, text="Remover E-mail", background="#A4A4A4",foreground="black", command=removeEmail) 
        deletButton.configure(font=("default", 10, "bold"))
        deletButton.place(x=90,y=90, width=115, height=20)
    except:
        NotCreateBanke = Label(windonws1, text='NENHUM BANCO DE DADOS ENCONTRADO', background="red",foreground="white", anchor=W)
        NotCreateBanke.configure(font=("default", 10, "bold"))
        NotCreateBanke.place(x=10,y=55)         

def criarBanco():
    windonws4 = Tk()
    windonws4.title('Criar Banco de dados')
    windonws4.geometry("322x130")
    windonws4.wm_resizable(width=False, height=False)
    windonws4.configure(background="#1C1C1C")

    menubar(windonws4)

    NomeBanco = Label(windonws4, text="Nome do Banco de Dados", background="#1C1C1C",foreground="white", anchor=W)
    NomeBanco.configure(font=("default", 10, "bold"))
    NomeBanco.place(x=8,y=12)

    NomeBanc = Entry(windonws4)
    NomeBanc.configure(font=("default", 10, "normal"))
    NomeBanc.place(x=10,y=35, width=300, height=20)

    def Create():
        caminho = f'C:/Users/Emails/{NomeBanc.get()}.csv'
        if os.path.exists(caminho):
            # os.makedirs("C:/Users/Emails")

            failBanco = Label(windonws4, text="Banco de dados existente", background="red",foreground="white")
            failBanco.configure(font=("default", 10, "normal"))
            failBanco.place(x=10,y=85, width=210, height=25)    
        else:
            with open(caminho.lower(), 'w', newline='') as file:
                writer = csv.writer(file, delimiter=';')
                dados = [';']
                writer.writerow(dados)

            sucessBanco = Label(windonws4, text="Banco de dados criado", background="green",foreground="white")
            sucessBanco.configure(font=("default", 10, "normal"))
            sucessBanco.place(x=10,y=85, width=210, height=25)    

    criarButton = Button(windonws4, text='Criar', background="#A4A4A4",foreground="black", command=Create)
    criarButton.configure(font=("default", 10, "bold"))
    criarButton.place(x=10,y=62, width=55,height=20) 

def sendEmail():
    windonws2 = Tk()
    windonws2.title("Send Email")
    windonws2.geometry("850x440")
    windonws2.wm_resizable(width=False, height=False)
    windonws2.configure(background="#1C1C1C")

    menubar(windonws2)
    
    Email = Label(windonws2, text="Usuário do E-mail", background="#1C1C1C",foreground="white", anchor=W)
    Email.configure(font=("default", 10, "bold"))
    Email.place(x=8,y=12)

    EmailAddress = Entry(windonws2)
    EmailAddress.configure(font=("default", 10, "normal"))
    EmailAddress.place(x=10,y=35, width=300, height=20)

    password3 = Label(windonws2, text='Senha do E-mail', background="#1C1C1C",foreground="white", anchor=W)
    password3.configure(font=("default", 10, "bold"))
    password3.place(x=8,y=62)

    passwd1 = Entry(windonws2, show="*")
    passwd1.configure(font=("default", 10, "bold"))
    passwd1.place(x=10,y=85,width=300,height=20)
    try:
        pasta = 'C:/Users/Emails'
        for diretorio, subpastas, arquivos in os.walk(pasta):
            arquivo = arquivos

        options_list = arquivo
        value_inside = tk.StringVar(windonws2)
        value_inside.set("Selecione o Banco de Dados")
        question_menu = tk.OptionMenu(windonws2, value_inside, *options_list)
        question_menu.place(x=9,y=115)
        
    except:
        NotCreateBanke = Label(windonws2, text='NENHUM BANCO DE DADOS ENCONTRADO', background="red",foreground="white", anchor=W)
        NotCreateBanke.configure(font=("default", 10, "bold"))
        NotCreateBanke.place(x=9,y=115)

    assuntoEmail = Label(windonws2, text='Assunto do E-mail', background="#1C1C1C",foreground="white", anchor=W)
    assuntoEmail.configure(font=("default", 10, "bold"))
    assuntoEmail.place(x=8,y=150)

    assunto = Entry(windonws2)
    assunto.configure(font=("default", 10, "normal"))
    assunto.place(x=10,y=173,width=300,height=20)

    corpoEmail = Label(windonws2, text='Corpo do E-mail', background="#1C1C1C",foreground="white", anchor=W)
    corpoEmail.configure(font=("default", 10, "bold"))
    corpoEmail.place(x=8,y=203)

    corpoText = tk.Frame(windonws2)
    corpoText.place(x=10, y=225)

    scrollbar = Scrollbar(corpoText)
    corpoTexto = tk.Text(corpoText, height=10, width=100, yscrollcommand=scrollbar.set)
    scrollbar.config(command=corpoTexto.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    corpoTexto.pack(side="left")
    corpoTexto.configure(font=("Arial", 11, "normal"))

    def send():
        try:
            with open(f'{pasta}/{value_inside.get()}', 'r') as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')

                for emails in csv_reader:
                    pass

            msg = EmailMessage()
            msg['Subject'] = assunto.get()
            msg['From'] = EmailAddress.get()
            msg['To'] = emails
            # corpo da mensagem
            msg.set_content(corpoTexto.get("1.0",'end-1c'))

            with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                smtp.login(EmailAddress.get(), passwd1.get())
                smtp.send_message(msg)
            emailSucess = Label(windonws2, text='[SUCESS] Email enviado', background="green",foreground="white")
            emailSucess.configure(font=("default", 10, "normal"))
            emailSucess.place(x=270,y=410, width=240, height=25)
        except:
            emailFail = Label(windonws2, text='[ERRO] Email não enviado', background="red",foreground="white")
            emailFail.configure(font=("default", 10, "normal"))
            emailFail.place(x=270,y=410, width=240, height=25)
    def sendAnexo():
        try:
            with open(f'{pasta}/{value_inside.get()}', 'r') as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')

                for emails in csv_reader:
                    pass

            msg = EmailMessage()
            msg['Subject'] = assunto.get()
            msg['From'] = EmailAddress.get()
            msg['To'] = emails
            # corpo da mensagem
            msg.set_content(corpoTexto.get("1.0",'end-1c'))
            caminhoArqs = filedialog.askopenfilenames()

            for caminhoArq in caminhoArqs:
                with open(caminhoArq, 'rb') as arquivos:
                    arquivo = arquivos.read()
                    extensao = os.path.splitext(caminhoArq)
                    extensaoTratado = str(extensao[1]).replace(".", "")
                    nameArq = f'''anexo{extensao[1]}'''
                    extensao1 = []
                    extensao1.append(extensaoTratado)
                    nameArq1 = []
                    nameArq1.append(nameArq)
                msg.add_attachment(arquivo, maintype='image', subtype=extensao1[-1], filename=nameArq1[-1])

            with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                smtp.login(EmailAddress.get(), passwd1.get())
                smtp.send_message(msg)
            emailSucess = Label(windonws2, text='[SUCESS] Email enviado', background="green",foreground="white")
            emailSucess.configure(font=("default", 10, "normal"))
            emailSucess.place(x=270,y=410, width=240, height=25)
        except:
            emailFail = Label(windonws2, text='[ERRO] Email não enviado', background="red",foreground="white")
            emailFail.configure(font=("default", 10, "normal"))
            emailFail.place(x=270,y=410, width=240, height=25)

    enviarButton = Button(windonws2, text="Enviar", background='#A4A4A4', foreground="black", command=send)
    enviarButton.configure(font=("default", 10, "bold"))
    enviarButton.place(x=10,y=410, width=100, height=25)

    enviarAnexoButton = Button(windonws2, text="Enviar com Anexo", background='#A4A4A4', foreground="black", command=sendAnexo)
    enviarAnexoButton.configure(font=("default", 10, "bold"))
    enviarAnexoButton.place(x=130,y=410, width=135, height=25)

entrarButton = Button(windonws, text="Entrar", background="#A4A4A4", foreground="black", command=userLogin)
entrarButton.configure(font=("default", 10, "bold"))
entrarButton.place(x=120, y=240, width=50, height=25)

windonws.mainloop()