import openpyxl as op
from openpyxl.styles import NamedStyle
import pandas as pd
import datetime as dt
from datetime import timezone
import time

# Pegando data de hoje
data_hoje = dt.datetime.now()
# Criando a classe para fornecedores / pedidos / emails


class Fornecedor():
    def __init__(self, Nome, Email):
        self.Nome = Nome
        self.Email = Email
        self.TotalPedidos = []

    def inserirPedidos(self, incremento_pedidos):
        self.TotalPedidos.append(incremento_pedidos)

    def mostrarPedidos(self):
        print(self.TotalPedidos)

    def removerPedido(self, Valor):
        print("------------------------")
        self.Valor = Valor
        pedido_removido = self.TotalPedidos[Valor]
        self.TotalPedidos.pop(Valor)
        print("Pedidos após a remoção-----------------")
        print(self.TotalPedidos)
        print("------------------------")


Pedidos = pd.read_excel('EntregasPendentes10_07_2023.xlsx')

Pedidos = Pedidos[Pedidos['Situação'] != 'Envio pendente']

Pedidos = Pedidos[Pedidos['Nacionalidade'] == 'Brasil']

valoresRateio = ['MATERIA-PRIMA',
                 'MATERIA PRIMA INDUSTRIALIZAÇÃO', 'MATERIAL DE USO E CONSUMO']
Pedidos = Pedidos[Pedidos['Rateio'].isin(valoresRateio)]

Pedidos.to_excel('PedidosAtraso.xlsx')

time.sleep(5)

tabelapd = pd.read_excel("./PedidosAtraso.xlsx")


# Puxando fornecedores sem duplicatas
fornecedores = tabelapd.loc[:, ['Fornecedor']].drop_duplicates(
    subset="Fornecedor", keep="first").values.tolist()

PedidosAMP = tabelapd.loc[tabelapd['Fornecedor'] == 'AMPHENOL TFC DO BRASIL LTDA', [
    "Neg.", "Data de entrega", "Cod.", "Material", "Faltam"]].reset_index()
PedidosAMP.pop(PedidosAMP.columns[0])

PedidosAMP.index += 1

PedidosAMP['Data de entrega'] = pd.to_datetime(
    PedidosAMP['Data de entrega'], format='%d/%m/%Y')

PedidosAMP['Data de entrega'] = PedidosAMP["Data de entrega"].dt.strftime(
    "%d/%m/%Y   ")

PedidosAMP.to_excel('PedidosAMP.xlsx')