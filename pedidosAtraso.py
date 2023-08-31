import pandas as pd
import datetime as dt
import time
import win32com.client as win32
# from tkinter import *
# from tkinter import Tk, Button, filedialog
import tkinter as tk
from tkinter import filedialog
import sendEmail

# Pegando data de hoje
data_hoje = dt.datetime.now()
# Criando a classe para fornecedores / pedidos / emails


class Fornecedor():
    def __init__(self, Nome, Email, TotalPedidos):
        self.Nome = Nome
        self.Email = Email
        self.TotalPedidos = TotalPedidos


def userInterface():
    janela = tk.Tk()
    janela.title("FollowUp F&K")
    janela.geometry("500x500")
    janela.resizable(False, False)
    string_path = tk.StringVar()
    string_path.set("Arquivo Selecionado")

    def add_file():
        global file_path
        file_path = filedialog.askopenfilenames()
        print('tuple', file_path)
        file_path = "".join(file_path)
        string_path.set(file_path)

    fileDialogButton = tk.Button(
        janela, text="Adicionar Arquivo", command=add_file)
    fileDialogButton.pack(pady=20)

    selectlabel = tk.Label(janela, textvariable=string_path)
    selectlabel.pack()

    send_emails = tk.Button(janela, text="Enviar Emails",
                            command=data_push_pandas)
    send_emails.pack(pady=10)

    janela.mainloop()


def formatar_dados(Pedidos):
    Pedidos.pop(Pedidos.columns[0])

    Pedidos.index += 1

    Pedidos['Data de entrega'] = pd.to_datetime(
        Pedidos['Data de entrega'], format='%d/%m/%Y')

    Pedidos['Data de entrega'] = Pedidos["Data de entrega"].dt.strftime(
        "%d/%m/%Y   ")

    # Caso queira criar arquivo excel
    # Pedidos.to_excel('Pedidos.xlsx')


def data_push_pandas():
    print("PATH DENTRO DO PANDAS", file_path)
    Pedidos = pd.read_excel(file_path)
    print(Pedidos)
    Pedidos = Pedidos[Pedidos['Situação'] != 'Envio pendente']

    Pedidos = Pedidos[Pedidos['Nacionalidade'] == 'Brasil']

    valoresRateio = ['MATERIA-PRIMA',
                     'MATERIA PRIMA INDUSTRIALIZAÇÃO', 'MATERIAL DE USO E CONSUMO']

    Pedidos = Pedidos[Pedidos['Rateio'].isin(valoresRateio)]

    # Pedidos.to_excel('PedidosAtraso.xlsx')
    Pedidos['Fornecedor'].to_string()


""" time.sleep(3)

tabelapd = pd.read_excel("./PedidosAtraso.xlsx")

tabelapd['Fornecedor'].to_string()
# Puxando fornecedores sem duplicatas
fornecedores = tabelapd.loc[:, ['Fornecedor']].drop_duplicates(
    subset="Fornecedor", keep="first").values.tolist()

# Pegando os pedidos de cada fornecedor e separando
Lista_fornecedores = []
for fornecedor in fornecedores:
    PedidosAtrasados = tabelapd.loc[tabelapd['Fornecedor'] == fornecedor[0], [
        "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()
    PedidosAtrasados.index.name = "N"
    formatar_dados(PedidosAtrasados)
    Lista_fornecedores.append(Fornecedor(
        fornecedor[0], f"{fornecedor[0]}@gmail.com", PedidosAtrasados))
    # Comando para gerar arquivos excel bom base nos pedidos e nomes de cada fornecedor
    # PedidosAtrasados.to_excel(f'Pedidos{fornecedor[0]}.xlsx')

outlook = win32.Dispatch('outlook.application') """

""" script =
<script>
document.getElementsByTagName('th').firstChild.text = 'N°'
</script> """

userInterface()
