import pandas as pd
import datetime as dt
import time
import win32com.client as win32

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


# Pegando os pedidos de cada fornecedor e separando

# Pegando os pedidos de cada fornecedor e separando
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

outlook = win32.Dispatch('outlook.application')

style = """
<style>
* {
padding: 5px;
}

thead {
  text-align: center;
  background-color: cadetblue;
}

tr, th,td {
  text-align: center;
  justify-content: center;
}

td:nth-child(5) {
  text-align: left;
  background-color: red;
}
</style>
"""

for fornc in Lista_fornecedores:
    print(F'Nome: {fornc.Nome}')
    print(F'Email: {fornc.Email}')
    print(F'Nome: {fornc.TotalPedidos}')
