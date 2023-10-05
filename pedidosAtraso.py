import pandas as pd
import datetime as dt
from datetime import timedelta
import time
import win32com.client as win32
import tkinter as tk
import customtkinter
from CTkListbox import *
import pygame

# Getting today date
today_date = dt.datetime.now()
iconpath = "C:/Users/Rafaeros/Documents/Development/Python/FollowUp/fk-logo.ico"


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

class interface():
    def __init__(self, master):

        #Declarating variables
        self.cDeletedSuppliers = []
        self.index = -5

        self.master = master
        master.title("Follow Up F&K Group")
        master.geometry("500x500")
        pygame.mixer.init()

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

        self.sendlateOrder_emails = customtkinter.CTkButton(master, text="Enviar Emails Atrasados", command=lambda m="corrective": self.clickevent(m))
        self.sendlateOrder_emails.pack(pady=10)

        self.sendPreventive_emails = customtkinter.CTkButton(
            master, text="Enviar Emails Preventivos", command=lambda m="preventive": self.clickevent(m))
        self.sendPreventive_emails.pack(pady=10)

    def add_email_file(self):
        global email_data_filepath
        email_data_filepath = customtkinter.filedialog.askopenfilename()
        email_data_filepath = "".join(email_data_filepath)
        if(email_data_filepath != ""):
            self.selectedArchive(email_data_filepath)
        else:
            self.selectedArchive(email_data_filepath)
            self.archiveTopLevel.destroy()

    def add_file(self):
        global orders_data_filepath
        orders_data_filepath = customtkinter.filedialog.askopenfilename()
        orders_data_filepath = "".join(orders_data_filepath)
        if(orders_data_filepath != ""):
            self.selectedArchive(orders_data_filepath)
        else:
            self.selectedArchive(orders_data_filepath)
            self.archiveTopLevel.destroy()

    
    def addPreventiveWindow(self):
        self.pTopLevel = ""
        self.icon.iconbitmap("./fk-log.ico")
        self.pTopLevel.window.grab_set()
        self.pTopLevel.window.mainloop()
    
    def addCorrectiveWindow(self):
        #Window configuration
        self.cTopLevel = customtkinter.CTkToplevel()
        self.cTopLevel.title("Enviando emails corretivos...")
        self.cTopLevel.geometry("800x600")
        self.cTopLevel.grab_set()
        self.cTopLevel.columnconfigure(0, weight=3)
        self.cTopLevel.columnconfigure(1, weight=3)
        self.cTopLevel.columnconfigure(2, weight=3)
        self.cTopLevel.rowconfigure(0, weight=5)
        self.cTopLevel.rowconfigure(1, weight=2)
        self.cTopLevel.rowconfigure(2, weight=3)
        self.cTopLevel.rowconfigure(3, weight=3)
        self.cTopLevel.rowconfigure(4, weight=3)

        self.cListBox = CTkListbox(self.cTopLevel, width=500, height=300)
        for correctiveSupplier_name in correctiveSuppliers_Names:
            self.cListBox.insert("END",correctiveSupplier_name)
        self.cListBox.grid(row=0, column=1, pady=10)

        self.restoreButton = customtkinter.CTkButton(self.cTopLevel, text="Restaurar", command=self.restoreListTopLevel)
        self.restoreButton.grid(row=1, column=0, pady=10, padx=10)
        
        self.deleteButton = customtkinter.CTkButton(self.cTopLevel, text="Deletar", command=self.deleteSelectedItem, fg_color="#FF0000", text_color="white", hover_color="#990000")
        self.deleteButton.grid(row=1, column=2, pady=10, padx=10)

        self.cancelButton = customtkinter.CTkButton(self.cTopLevel,text="Cancelar", command=self.cTopLevel.destroy, width=300, height=50)
        self.cancelButton.grid(row=3, column=0, pady=40, padx=40)
        self.sendButton = customtkinter.CTkButton(self.cTopLevel, text="Enviar Email", command=lambda: sendCorrectiveEmail(correctiveSuppliers_List), width=300, height=50)
        self.sendButton.grid(row=3, column=2, pady=40, padx=40)
        
    def clickevent(self, click):
        global sendChoose
        sendChoose = click
        data_push()
        if(sendChoose=="corrective"):
            self.addCorrectiveWindow()
        elif(sendChoose=="preventive"):
            self.addPreventiveWindow()
    
    def selectedArchive(self, path):
        self.archiveTopLevel = customtkinter.CTkToplevel()
        self.archiveTopLevel.title("Arquivo selecionado")
        self.archiveTopLevel.geometry("300x200")
        self.archiveTopLevel.grab_set()

        self.playNotificationSound()

        #Shows "Selected Arqhive" in the window
        self.selectedArchiveLabel = customtkinter.CTkLabel(self.archiveTopLevel, text="Arquivo selecionado:", pady=10,padx=10)
        self.selectedArchiveLabel.pack(pady=10,padx=10)

        #splits the file path
        splitFilePath = path.split('/')
        splitLen = len(splitFilePath)-1
        text=splitFilePath[splitLen]
        text = customtkinter.StringVar()
        text.set(f"{splitFilePath[splitLen]}")
        
        #underline text configuration
        underlineText = customtkinter.CTkFont(underline=True)

        #Archive name label show
        self.selectedArchiveNameLabel = customtkinter.CTkLabel(self.archiveTopLevel, font=underlineText, textvariable=text)
        self.selectedArchiveNameLabel.pack(pady=10, padx=10)

        self.okButton = customtkinter.CTkButton(self.archiveTopLevel, text="OK", command=self.archiveTopLevel.destroy)
        self.okButton.pack(pady=20, padx=20)

    def deleteSelectedItem(self):
        self.index = self.cListBox.curselection()
        self.cListBox.delete(self.index)

        for supplier in correctiveSuppliers_List:
            if(supplier.Name==correctiveSuppliers_List[self.index].Name):
                self.cDeletedSuppliers.append(correctiveSuppliers_List[self.index])
                break

        
        print("Fornecedor deletado")
        print(correctiveSuppliers_List[self.index].Name)

        correctiveSuppliers_List.pop(self.index)

        print("Lista após o delete")
        for supplier in correctiveSuppliers_List:
            print(supplier.Name)
        print(f"Tamanho lista: {len(correctiveSuppliers_List)}")

    def restoreListTopLevel(self):
        lastDeletedSupplier = len(self.cDeletedSuppliers)
        if(lastDeletedSupplier==0):
            self.emptyListTopLevel = customtkinter.CTkToplevel()
            self.emptyListTopLevel.title("Erro")
            self.emptyListTopLevel.geometry("300x200")
            self.emptyListTopLevel.grab_set()
            self.emptyListLabel = customtkinter.CTkLabel(self.emptyListTopLevel, text="Nenhum fornecedor foi deletado anteriormente")
            self.emptyListLabel.pack(pady=10, padx=10)
            self.emptyListButton = customtkinter.CTkButton(self.emptyListTopLevel, text="OK", command=self.emptyListTopLevel.destroy)
            self.emptyListButton.pack(pady=10, padx=10)

        elif(lastDeletedSupplier>0):
            self.deletedListTopLevel = customtkinter.CTkToplevel()
            self.deletedListTopLevel.title("Index")
            self.deletedListTopLevel.geometry("300x300")
            self.deletedListTopLevel.grab_set()

            self.ctkIndexList = CTkListbox(self.deletedListTopLevel)
            for fornecedor in self.cDeletedSuppliers:
                self.ctkIndexList.insert("END",f"{fornecedor.Name}")
            self.ctkIndexList.pack()

            self.button = customtkinter.CTkButton(self.deletedListTopLevel, width=100, height=100, text_color="RED", text="OK", command=self.restoreListCommand)
            self.button.pack(pady=10) #to aqui

    def restoreListCommand(self):
            self.index = self.ctkIndexList.curselection()
            self.ctkIndexList.delete(self.index)
            
            self.cListBox.insert("END", self.cDeletedSuppliers[self.index].Name)

            for fornecedor in self.cDeletedSuppliers:
                if(fornecedor.Name == self.cDeletedSuppliers[self.index].Name):
                    correctiveSuppliers_List.append(self.cDeletedSuppliers[self.index])
                    print(f"Fornecedor restaurado: {fornecedor.Name}")
                    self.cDeletedSuppliers.pop(self.index)
                    print(F"Tamanho da lista dos deletados após restaurar: {len(self.cDeletedSuppliers)}")
                    break
            if(self.cDeletedSuppliers==[]):
                self.deletedListTopLevel.destroy()

    def playNotificationSound(self):
        pygame.mixer.music.load('./Notify.wav')
        pygame.mixer.music.play(loops=0)

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

        for supplier in correctiveSuppliers_Names:
            print(supplier)

        correctiveSuppliers_List = []
        for correctiveSupplier_Name in correctiveSuppliers_Names:
            lateOrders = total_late_orders.loc[total_late_orders['Fornecedor'] == correctiveSupplier_Name[0], [
                "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()
            format_data(lateOrders)

            cCurrent_email = emails_data.loc[emails_data['Nome'] == correctiveSupplier_Name[0], [
                "Email"]]
            correctiveSuppliers_List.append(
                Supplier(correctiveSupplier_Name[0], f"{cCurrent_email}", lateOrders))
            # Supplier(Name, Email, Totalorders, Index)

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

            preventiveSuppliers_List.append(
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
root.iconbitmap(iconpath)
userinterface = interface(root)
root.mainloop()
