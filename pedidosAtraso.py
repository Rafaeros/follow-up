import pandas as pd
import datetime as dt
import time


# Pegando data de hoje
data_hoje = dt.datetime.now()
# Criando a classe para fornecedores / pedidos / emails


def formatar_dados(Pedidos):
    Pedidos.pop(Pedidos.columns[0])

    Pedidos.index += 1

    Pedidos['Data de entrega'] = pd.to_datetime(
        Pedidos['Data de entrega'], format='%d/%m/%Y')

    Pedidos['Data de entrega'] = Pedidos["Data de entrega"].dt.strftime(
        "%d/%m/%Y   ")

    # Caso queira criar arquivo excel
    # Pedidos.to_excel('Pedidos.xlsx')


class Fornecedor():
    def __init__(self, Nome, Email, TotalPedidos):
        self.Nome = Nome
        self.Email = Email
        self.TotalPedidos = TotalPedidos

    def incrementarFornecedor(self, forn):
        self.Nome.append(forn)
    """ def inserirPedidos(self, incremento_pedidos):
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
        print("------------------------") """


Pedidos = pd.read_excel('EntregasPendentes10_07_2023.xlsx')

Pedidos = Pedidos[Pedidos['Situação'] != 'Envio pendente']

Pedidos = Pedidos[Pedidos['Nacionalidade'] == 'Brasil']

valoresRateio = ['MATERIA-PRIMA',
                 'MATERIA PRIMA INDUSTRIALIZAÇÃO', 'MATERIAL DE USO E CONSUMO']

Pedidos = Pedidos[Pedidos['Rateio'].isin(valoresRateio)]

Pedidos.to_excel('PedidosAtraso.xlsx')

time.sleep(3)

tabelapd = pd.read_excel("./PedidosAtraso.xlsx")

tabelapd['Fornecedor'].to_string()
# Puxando fornecedores sem duplicatas
fornecedores = tabelapd.loc[:, ['Fornecedor']].drop_duplicates(
    subset="Fornecedor", keep="first").values.tolist()

# Tentando iterar os pedidos

""" i = 0

while i < len(fornecedores):
    print(type(fornecedores[i]), fornecedores[i])
    i += 1 """
Lista_fornecedores = []
for fornecedor in fornecedores:
    PedidosAtrasados = tabelapd.loc[tabelapd['Fornecedor'] == fornecedor[0], [
        "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()
    formatar_dados(PedidosAtrasados)
    Lista_fornecedores.append(Fornecedor(
        fornecedor[0], f"{fornecedor[0]}@gmail.com", PedidosAtrasados))

    # Comando para gerar arquivos excel bom base nos pedidos e nomes de cada fornecedor
    # PedidosAtrasados.to_excel(f'Pedidos{fornecedor[0]}.xlsx')


# ---Printar no console os dados de cada fornecedor da classe Fornecedor
# Lista_fornecedores.append(Fornecedor(fornecedor, "Teste@gmail.com", pedidosFornecedor))
for fornc in Lista_fornecedores:
    print(F'Nome: {fornc.Nome}')
    print(F'Email: {fornc.Email}')
    print(F'Nome: {fornc.TotalPedidos}')
