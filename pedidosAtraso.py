import pandas as pd
import datetime as dt
import time
import win32com.client as win32
# from tkinter import *
# from tkinter import Tk, Button, filedialog
import tkinter as tk
from tkinter import filedialog
import chardet


# Pegando data de hoje
data_hoje = dt.datetime.now()
data_hoje = data_hoje.strftime('%d/%m/%Y')
print(type(data_hoje))
print(data_hoje)
# Criando a classe para fornecedores / pedidos / emails


class Fornecedor():
    def __init__(self, Nome, Email, TotalPedidos):
        self.Nome = Nome
        self.Email = Email
        self.TotalPedidos = TotalPedidos


def userInterface():
    janela = tk.Tk()
    janela.title("FollowUp F&K")
    janela.geometry("350x400")
    janela.resizable(False, False)
    string_path = tk.StringVar()
    string_path.set("Arquivo Selecionado")

    def add_email_file():
        global email_file_path
        email_file_path = filedialog.askopenfilenames()
        print(email_file_path)
        email_file_path = "".join(email_file_path)

    def add_file():
        global file_path
        file_path = filedialog.askopenfilenames()
        print(file_path)
        file_path = "".join(file_path)
        string_path.set(file_path)

    step_1 = tk.Label(janela, text="1° Passo")
    step_1.pack(pady=10)

    emailDialogButton = tk.Button(
        janela, text="Adicionar Arquivo C/ Emails", command=add_email_file)
    emailDialogButton.pack(pady=10)

    step_2 = tk.Label(janela, text="2° Passo")
    step_2.pack(pady=10)

    fileDialogButton = tk.Button(
        janela, text="Adicionar Arquivo C/ Pedidos", command=add_file)
    fileDialogButton.pack(pady=10)

    selectlabel = tk.Label(janela, textvariable=string_path)
    selectlabel

    step_3 = tk.Label(janela, text="3° Passo")
    step_3.pack(pady=30)

    send_emails = tk.Button(janela, text="Enviar Emails",
                            command=data_push_pandas)
    send_emails.pack(pady=10)

    janela.mainloop()


def formatar_dados(Orders):
    Orders.pop(Orders.columns[0])

    Orders.index += 1

    Orders['Data de entrega'] = pd.to_datetime(
        Orders['Data de entrega'], format='%d/%m/%Y')

    Orders['Data de entrega'] = Orders["Data de entrega"].dt.strftime(
        "%d/%m/%Y   ")


def data_push_pandas():

    suppliers_data = pd.read_excel(email_file_path)
    email_data = suppliers_data[["Nome", "Email"]]

    Pedidos = pd.read_excel(file_path)
    Pedidos = Pedidos[Pedidos['Situação'] != 'Envio pendente']
    Pedidos = Pedidos[Pedidos['Nacionalidade'] == 'Brasil']
    valoresRateio = ['MATERIA-PRIMA',
                     'MATERIA PRIMA INDUSTRIALIZAÇÃO', 'MATERIAL DE USO E CONSUMO']
    Pedidos = Pedidos[Pedidos['Rateio'].isin(valoresRateio)]
    Pedidos = Pedidos[Pedidos['Data de entrega'] < data_hoje]

    print(Pedidos)

    Pedidos['Fornecedor'].to_string()
    fornecedores = Pedidos.loc[:, ['Fornecedor']].drop_duplicates(
        subset="Fornecedor", keep="first").values.tolist()

    lista_fornecedores = []
    for fornecedor in fornecedores:
        lateOrders = Pedidos.loc[Pedidos['Fornecedor'] == fornecedor[0], [
            "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()
        lateOrders.index.name = "N"
        formatar_dados(lateOrders)

        current_email = email_data.loc[email_data['Nome']
                                       == fornecedor[0], ["Email"]]

        print(current_email)

        lista_fornecedores.append(Fornecedor(
            fornecedor[0], f"", lateOrders))


    # Comando para gerar arquivos excel bom base nos pedidos e nomes de cada fornecedor
    # PedidosAtrasados.to_excel(f'Pedidos{fornecedor[0]}.xlsx')
userInterface()
