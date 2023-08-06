import openpyxl as op
import pandas as pd
import datetime as dt

# Pegando data de hoje
data_hoje = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

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


# Carregando Planilha
ws = op.load_workbook('./EntregasPendentes10_07_2023.xlsx')
planilha_ativa = ws.active
ult_linha_planilha = planilha_ativa.max_row

# Vendo todos pedidos que estão atrasados
""" def pegandoPedidosAtrasados(database):
    i=0
    fill_atrasado = Pattern
    ultima_linha = database.max_row
    datas_antigas = [planilha_ativa.cell(row=i, column=2).value for i in range(2,ultima_linha):]
    
    if(planilha_ativa.cell(row=i, column=2)<data_hoje: planilha_ativa[f"B.{celula.row}"].fill = PatternFill(start_color) """


# Limpando planilha
for celula in planilha_ativa["Q"]:
    if (celula.value == "Envio pendente"):
        linha_celula = celula.row
        planilha_ativa.delete_rows(linha_celula)

for celula in planilha_ativa["G"]:
    if (celula.value != "Brasil" and celula.value != "Nacionalidade"):
        linha_celula = celula.row
        planilha_ativa.delete_rows(linha_celula)

for celula in planilha_ativa['M']:
    if (celula.value != "MATERIA-PRIMA" and celula.value != "Rateio"):
        linha_celula = celula.row
        planilha_ativa.delete_rows(linha_celula)


ws.save("PedidoAtraso.xlsx")

tabelapd = pd.read_excel("./PedidoAtraso.xlsx")

#Puxando fornecedores sem duplicatas
fornecedores = tabelapd.loc[:, ['Fornecedor']].drop_duplicates(subset="Fornecedor", keep="first").reset_index().values.tolist()
print(fornecedores[2])
print("total fornecedores: ", len(fornecedores))

for fornecedor in fornecedores:
    lateOrders = tabelapd.loc[tabelapd['Fornecedor']==f'{fornecedor}']
    print(lateOrders)



""" totalpedidosamp = tabelapd.loc[tabelapd["Fornecedor"]== "AMPHENOL TFC DO BRASIL LTDA"].values

print(totalpedidosamp)
print(len(totalpedidosamp)) """

# totalpedidosamp = totalpedidosamp.reset_index()


#classeFornecedorAmphenol = Fornecedor("Amphenol", "Amphenol.com.br")


#print(classeFornecedorAmphenol.Nome)
#print(classeFornecedorAmphenol.Email)

# print(len(classeFornecedorAmphenol.TotalPedidos))

""" classeFornecedorAmphenol.mostrarPedidos()
classeFornecedorAmphenol.removerPedido(0) """
