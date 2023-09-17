import pandas as pd
import datetime as dt
from datetime import timedelta
import time
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog


# Getting today date
today_date = dt.datetime.now()

# Email style
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


class Supplier():
    def __init__(self, Name, Email, TotalOrders):
        self.Name = Name
        self.Email = Email
        self.TotalOrders = TotalOrders


def userInterface():
    window = tk.Tk()
    window.title("FollowUp F&K")
    window.geometry("350x400")
    window.resizable(False, False)
    string_path = tk.StringVar()
    string_path.set("Arquivo Selecionado")

    def add_email_file():
        global email_data_filepath
        email_data_filepath = filedialog.askopenfilenames()
        print(email_data_filepath)
        email_data_filepath = "".join(email_data_filepath)

    def add_file():
        global orders_data_filepath
        orders_data_filepath = filedialog.askopenfilenames()
        print(orders_data_filepath)
        orders_data_filepath = "".join(orders_data_filepath)
        string_path.set(orders_data_filepath)

    step_1 = tk.Label(window, text="1° Passo")
    step_1.pack(pady=10)

    emailDialogButton = tk.Button(
        window, text="Adicionar Arquivo C/ Emails", command=add_email_file)
    emailDialogButton.pack(pady=10)

    step_2 = tk.Label(window, text="2° Passo")
    step_2.pack(pady=10)

    fileDialogButton = tk.Button(
        window, text="Adicionar Arquivo C/ Pedidos", command=add_file)
    fileDialogButton.pack(pady=10)

    selectlabel = tk.Label(window, textvariable=string_path)
    selectlabel

    step_3 = tk.Label(window, text="3° Passo")
    step_3.pack(pady=30)

    send_emails = tk.Button(window, text="Enviar Emails",
                            command=data_push)
    send_emails.pack(pady=10)

    window.mainloop()


def format_data(Orders):
    Orders.pop(Orders.columns[0])

    Orders.index += 1

    Orders['Data de entrega'] = pd.to_datetime(
        Orders['Data de entrega'], format='%d/%m/%Y')

    Orders['Data de entrega'] = Orders["Data de entrega"].dt.strftime(
        "%d/%m/%Y   ")


def data_push():
    suppliers_data = pd.read_excel(email_data_filepath)
    emails_data = suppliers_data[["Nome", "Email"]]

    total_orders = pd.read_excel(orders_data_filepath)
    total_orders = total_orders[total_orders['Situação'] != 'Envio pendente']
    total_orders = total_orders[total_orders['Nacionalidade'] == 'Brasil']
    MP_filter = ['MATERIA-PRIMA', 'MATERIA PRIMA INDUSTRIALIZAÇÃO',
                 'MATERIAL DE USO E CONSUMO']
    total_orders = total_orders[total_orders['Rateio'].isin(MP_filter)]

    total_orders['Data de entrega'] = pd.to_datetime(
        total_orders['Data de entrega'], format='%d/%m/%Y')

    # Late Orders for corrective treatment
    total_late_orders = total_orders[total_orders['Data de entrega'] < today_date]

    # Ten days ahead Orders for preventive preventive treatment
    date_tenDaysAhead = today_date + timedelta(days=11)
    dateMask = (total_orders['Data de entrega'] > today_date) & (
        total_orders['Data de entrega'] <= date_tenDaysAhead)
    orders_tenDaysAhead = total_orders.loc[dateMask]

    if (total_orders['Data de entrega'] < today_date):
        total_late_orders['Fornecedor'].to_string()

        suppliers = total_late_orders.loc[:, ['Fornecedor']].drop_duplicates(
            subset="Fornecedor", keep="first").values.tolist()

        suppliers_list = []
        for supplier in suppliers:
            lateOrders = total_late_orders.loc[total_late_orders['Fornecedor'] == supplier[0], [
                "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()
            lateOrders.index.name = "N"
            format_data(lateOrders)

            current_email = emails_data.loc[emails_data['Nome'] == supplier[0], [
                "Email"]]
            suppliers_list.append(
                Supplier(supplier[0], f"{current_email}", lateOrders))
            # Supplier(Name, Email, Totalorders)

        sendEmail(suppliers_list)

    # Comando para gerar arquivos excel bom base nos total_late_orders e nomes de cada fornecedor
    # PedidosAtrasados.to_excel(f'total_late_orders{fornecedor[0]}.xlsx')


def sendEmail(suppliersList):
    outlook = win32.Dispatch("Outlook.Application")
    for supplier in suppliersList:
        lateOrdersHTML = supplier.TotalOrders.to_html(
            col_space=50, justify='center')
        html_body = f"""
        <!DOCTYPE html>
        <html>
        <head>
            {style}
        </head>
        <body>
            <h1>Olá:{supplier.Nome}</h1>
            <h2>Favor validar esses pedidos que constam em atraso em nosso sistema: </h2>
            {lateOrdersHTML}
        </body>
        </html>
        """
        print(html_body)
        email = outlook.CreateItem(0)
        time.sleep(1)
        email.To = 'rafaelzinhobr159@gmail.com'
        email.Subject = f"Pedidos atrasados {supplier.Name}"
        email.HTMLBody = (html_body)
        email.Send()
        print(f"Email enviado: {supplier.Name}")
        time.sleep(2)


userInterface()
