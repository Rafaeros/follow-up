import pandas as pd
import datetime as dt
from datetime import timedelta
import time
import win32com.client as win32
import tkinter as tk
import customtkinter
from CTkListbox import *

# Getting today date
today_date = dt.datetime.now()
iconpath = "./fk-logo.ico"

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

class cTopLevel():
    def __init__(self):
        self.window = customtkinter.CTkToplevel()
        self.window.title("Enviando emails corretivos...")
        self.window.geometry("800x600")
        self.window.columnconfigure(0,weight=3)
        self.window.columnconfigure(1,weight=3)
        self.window.columnconfigure(2,weight=3)
        self.window.rowconfigure(0, weight=5)
        self.window.rowconfigure(1, weight=2)
        self.window.rowconfigure(2, weight=3)

        self.pListBox = CTkListbox(self.window,width=500, height=300)
        for correctiveSupplier_name in correctiveSuppliers_Names:
            self.pListBox.insert("END",f"{correctiveSupplier_name}")
        self.pListBox.grid(column=1, row=0, pady=20)

        self.deleteButton = customtkinter.CTkButton(self.window, text="Deletar", command=self.deleteSelectedItem, fg_color="#FF0000", text_color="white", hover_color="#990000")
        self.deleteButton.grid(row=1, column=1, pady=10, padx=10)

        self.cancelButton = customtkinter.CTkButton(self.window,text="Cancelar", command=self.window.destroy, width=300, height=50)
        self.cancelButton.grid(row=3, column=0, pady=40, padx=40)
        self.sendButton = customtkinter.CTkButton(self.window, text="Enviar Email", command=lambda: sendCorrectiveEmail(correctiveSuppliers_List), width=300, height=50)
        self.sendButton.grid(row=3, column=2, pady=40, padx=40)

    def deleteSelectedItem(self):
        index = self.pListBox.curselection()
        self.pListBox.delete(index)
        correctiveSuppliers_List.pop(index)
        for name in correctiveSuppliers_List:
            print(name.Name)

class pTopLevel():
    def __init__(self):
        self.title("Enviando emails preventivos...")
        self.window.geometry("600x600")

        cListBox = CTkListbox(self.window,width=500)
        for preventiveSuppliers_Name in preventiveSuppliers_Names:
            cListBox.insert("END",f"{preventiveSuppliers_Name}")
        cListBox.pack(pady=0)

class interface():
    def __init__(self, master):
        self.master = master
        master.title("Follow Up F&K Group")
        master.geometry("500x500")
        master.iconbitmap(iconpath)

        self.appearance = customtkinter.set_appearance_mode("Dark")
        self.theme = customtkinter.set_default_color_theme("dark-blue")

        self.step_1 = customtkinter.CTkLabel(master, text="1° Passo")
        self.step_1.pack(pady=30)

        self.emailDialogButton = customtkinter.CTkButton(
            master, text="Adicionar Arquivo C/ Emails", command=self.add_email_file)
        self.emailDialogButton.pack(pady=10)

        self.step_2 = customtkinter.CTkLabel(master, text="2° Passo")
        self.step_2.pack(pady=10)

        self.fileDialogButton = customtkinter.CTkButton(
            master, text="Adicionar Arquivo C/ Pedidos", command=self.add_file)
        self.fileDialogButton.pack(pady=10)

        self.step_3 = customtkinter.CTkLabel(master, text="3° Passo")
        self.step_3.pack(pady=10)

        self.sendlateOrder_emails = customtkinter.CTkButton(master, text="Enviar Emails Atrasados",
                                        command=lambda m="corrective": self.clickevent(m))
        self.sendlateOrder_emails.pack(pady=10)

        self.sendPreventive_emails = customtkinter.CTkButton(
            master, text="Enviar Emails Preventivos", command=lambda m="preventive": self.clickevent(m))
        self.sendPreventive_emails.pack(pady=10)

    def add_email_file(self):
        global email_data_filepath
        email_data_filepath = customtkinter.filedialog.askopenfilename()
        email_data_filepath = "".join(email_data_filepath)

    def add_file(self):
        global orders_data_filepath
        orders_data_filepath = customtkinter.filedialog.askopenfilename()
        orders_data_filepath = "".join(orders_data_filepath)
    
    def addPreventiveWindow(self):
        self.pTopLevel = ""
        self.icon.iconbitmap("./fk-log.ico")
        self.pTopLevel.window.grab_set()
        self.pTopLevel.window.mainloop()
    
    def addCorrectiveWindow(self):
        self.cTopLevel = cTopLevel()
        self.cTopLevel.window.grab_set()
        self.cTopLevel.window.mainloop()


    def clickevent(self, click):
        global sendChoose
        sendChoose = click
        data_push()
        if(sendChoose=="corrective"):
            self.addCorrectiveWindow()
        elif(sendChoose=="preventive"):
            self.addPreventiveWindow()

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

    if (sendChoose == "corrective"):
        print("Enviando emails atrasados")

        global correctiveSuppliers_Names
        global correctiveSuppliers_List

        correctiveSuppliers_Names = total_late_orders.loc[:, ['Fornecedor']].drop_duplicates(
            subset="Fornecedor", keep="first").values.tolist()

        correctiveSuppliers_List = []
        for correctiveSupplier_Name in correctiveSuppliers_Names:
            lateOrders = total_late_orders.loc[total_late_orders['Fornecedor'] == correctiveSupplier_Name[0], [
                "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()
            format_data(lateOrders)

            pCurrent_email = emails_data.loc[emails_data['Nome'] == correctiveSupplier_Name[0], [
                "Email"]]
            correctiveSuppliers_List.append(
                Supplier(correctiveSupplier_Name[0], f"{pCurrent_email}", lateOrders))
            # Supplier(Name, Email, Totalorders)

            for names in correctiveSuppliers_List:
                print(names.Name)
                print(names.Email)
                print(names.TotalOrders)

    elif (sendChoose == "preventive"):
        global preventiveSuppliers_Names
        global preventiveSuppliers_List

        preventiveSuppliers_Names = orders_tenDaysAhead.loc[:, ['Fornecedor']].drop_duplicates(
            subset="Fornecedor", keep="first").values.tolist()
        
        preventiveSuppliers_List = []
        for preventiveSupplier_Name in preventiveSuppliers_Names:
            preventiveOrders = orders_tenDaysAhead.loc[orders_tenDaysAhead['Fornecedor'] == correctiveSupplier_Name[0], [
                "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]]
            preventiveOrders.index.name = "N"
            format_data(preventiveOrders)

            cCurrent_email = emails_data.loc[emails_data['Nome'] == correctiveSupplier_Name[0], ["Email"]]

            correctiveSuppliers_List.append(
                Supplier(correctiveSupplier_Name[0], cCurrent_email, preventiveOrders))
            #Class Supplier(Name, Email, Orders)

        # Comando para gerar arquivos excel bom base nos total_late_orders e nomes de cada fornecedor
        # PedidosAtrasados.to_excel(f'total_late_orders{fornecedor[0]}.xlsx')

def sendCorrectiveEmail(suppliersList):
    outlook = win32.Dispatch("Outlook.Application")
    ccEmail = ["glaucio.costa@fkgroup.com.br","luciana.santos@fkgroup.com.br", "guilherme.silva@fkgroup.com.br"]
    time.sleep(1)
    for supplier in suppliersList:
        lateOrdersHTML = supplier.TotalOrders.to_html(
            col_space=50, justify='center')
        correctiveEmailBody = f"""
        <!DOCTYPE html>
        <html>
        <head>
            {style}
        </head>
        <body>
            <h1>Olá:{supplier.Name}</h1>
            <h2>Favor validar esses pedidos que constam em atraso em nosso sistema: </h2>
            {lateOrdersHTML}
        </body>
        </html>
        """
        print(correctiveEmailBody)
        email = outlook.CreateItem(0)
        time.sleep(1)
        email.To = f'{supplier.Email}'
        email.Cc = ccEmail
        email.Subject = f"Pedidos atrasados {supplier.Name}"
        email.HTMLBody = (correctiveEmailBody)
        email.Send()
        print(f"Email enviado: {supplier.Name}")
        time.sleep(2)

def sendPreventiveEmail(suppliersList):
    outlook = win32.Dispatch("Outlook.Application")
    time.sleep(1)
    for supplier in suppliersList:
        lateOrdersHTML = supplier.TotalOrders.to_html(
            col_space=50, justify='center')
        preventiveEmailBody = f"""
        <!DOCTYPE html>
        <html>
        <head>
            {style}
        </head>
        <body>
            <h1>Olá:{supplier.Name}</h1>
            <h2>Favor validar confirmar a entrega desses pedidos conforme as datas previstas: </h2>
            {lateOrdersHTML}
        </body>
        </html>
        """
        print(preventiveEmailBody)
        email = outlook.CreateItem(0)
        time.sleep(1)
        email.To = 'rafaelzinhobr159@gmail.com'
        email.Subject = f"Entrega Pedidos: {supplier.Name}"
        email.HTMLBody = (preventiveEmailBody)
        email.Send()
        print(f"Email enviado: {supplier.Name}")
        time.sleep(2)

root = customtkinter.CTk()
userinterface = interface(root)
root.mainloop()