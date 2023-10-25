import pandas as pd
import datetime as dt
from datetime import timedelta
import time
import win32com.client as win32
import tkinter as tk
import customtkinter as ctk
from CTkListbox import *
import pygame

# Getting today date
today_date = dt.datetime.now()
iconpath = "fk-logo.ico"

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
        self.emailCcList = []
        self.index = -5

        self.master = master
        master.title("Follow Up F&K Group")
        master.geometry("500x500")
        pygame.mixer.init()

        self.appearance = ctk.set_appearance_mode("Dark")
        self.theme = ctk.set_default_color_theme("dark-blue")

        self.step_1 = ctk.CTkLabel(master, text="1° Passo")
        self.step_1.pack(pady=30)

        self.emailDialogButton = ctk.CTkButton(
            master, text="Adicionar Arquivo C/ Emails", command=self.add_email_file)
        self.emailDialogButton.pack(pady=10)

        self.step_2 = ctk.CTkLabel(master, text="2° Passo")
        self.step_2.pack(pady=10)

        self.fileDialogButton = ctk.CTkButton(
            master, text="Adicionar Arquivo C/ Pedidos", command=self.add_file)
        self.fileDialogButton.pack(pady=10)

        self.step_3 = ctk.CTkLabel(master, text="3° Passo")
        self.step_3.pack(pady=10)

        self.sendlateOrder_emails = ctk.CTkButton(master, text="Enviar Emails Atrasados", command=lambda m="corrective": self.clickevent(m))
        self.sendlateOrder_emails.pack(pady=10)

        self.sendPreventive_emails = ctk.CTkButton(
            master, text="Enviar Emails Preventivos", command=lambda m="preventive": self.clickevent(m))
        self.sendPreventive_emails.pack(pady=10)

    def add_email_file(self):
        global email_data_filepath
        email_data_filepath = ctk.filedialog.askopenfilename()
        email_data_filepath = "".join(email_data_filepath)
        self.selectedArchive(email_data_filepath)

    def add_file(self):
        global orders_data_filepath
        orders_data_filepath = ctk.filedialog.askopenfilename()
        orders_data_filepath = "".join(orders_data_filepath)
        self.selectedArchive(orders_data_filepath)

    def format_data(self, Orders):
        Orders.pop(Orders.columns[0])

        Orders.index += 1

        Orders['Data de entrega'] = pd.to_datetime(
            Orders['Data de entrega'], format='%d/%m/%Y')

        Orders['Data de entrega'] = Orders["Data de entrega"].dt.strftime(
            "%d/%m/%Y   ")

    def data_push(self):
        suppliersData = pd.read_excel(email_data_filepath)
        emails_data = suppliersData[["Nome", "Email"]]

        total_orders = pd.read_excel(orders_data_filepath)
        total_orders = total_orders[total_orders['Situação'] != 'Envio pendente']
        total_orders = total_orders[total_orders['Nacionalidade'] == 'Brasil']
        MP_filter = ['MATERIA-PRIMA', 'MATERIA PRIMA INDUSTRIALIZAÇÃO',
                    'MATERIAL DE USO E CONSUMO']
        total_orders = total_orders[total_orders['Rateio'].isin(MP_filter)]

        total_orders['Data de entrega'] = pd.to_datetime(
            total_orders['Data de entrega'], format='%d/%m/%Y')

        # Late Orders for corrective treatment
        lastDay = today_date - timedelta(days=1)
        total_late_orders = total_orders[total_orders['Data de entrega'] < lastDay]

        # Ten days ahead Orders for preventive preventive treatment
        date_tenDaysAhead = today_date + timedelta(days=11)
        dateMask = (total_orders['Data de entrega'] > today_date) & (
            total_orders['Data de entrega'] <= date_tenDaysAhead)
        orders_tenDaysAhead = total_orders.loc[dateMask]

        if (sendChoose == "corrective"):
            print("Enviando emails atrasados")
            global correctiveSuppliersNamesList
            global correctiveSuppliersNames
            global correctiveSuppliersList

            correctiveSuppliersNames = []

            correctiveSuppliersNamesList = total_late_orders.loc[:, ['Fornecedor']].drop_duplicates(
                subset="Fornecedor", keep="first").values.tolist()
            
            for Name in correctiveSuppliersNamesList:
                correctiveSuppliersNames.append(Name[0])

            for supplier in correctiveSuppliersNames:
                print(supplier)

            correctiveSuppliersList = []
            for Name in correctiveSuppliersNames:
                lateOrders = total_late_orders.loc[total_late_orders['Fornecedor'] == Name, [
                    "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()
                self.format_data(lateOrders)

                cCurrent_email = emails_data.loc[emails_data['Nome'] == Name, [
                    "Email"]]
                correctiveSuppliersList.append(
                    Supplier(Name, f"{cCurrent_email}", lateOrders))
                # Supplier(Name, Email, Totalorders, Index)

        elif (sendChoose == "preventive"):
            global preventiveSuppliers_Names
            global preventiveSuppliers_List

            preventiveSuppliers_Names = orders_tenDaysAhead.loc[:, ['Fornecedor']].drop_duplicates(
                subset="Fornecedor", keep="first").values.tolist()
            
            preventiveSuppliers_List = []
            for preventiveSupplier_Name in preventiveSuppliers_Names:
                preventiveOrders = orders_tenDaysAhead.loc[orders_tenDaysAhead['Fornecedor'] == preventiveSupplier_Name[0], [
                    "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]]
                preventiveOrders.index.name = "N"
                self.format_data(preventiveOrders)

                pCurrent_email = emails_data.loc[emails_data['Nome'] == preventiveSupplier_Name[0], ["Email"]]

                preventiveSuppliers_List.append(
                    Supplier(preventiveSupplier_Name[0], pCurrent_email, preventiveOrders))
                #Class Supplier(Name, Email, Orders)

            # Comando para gerar arquivos excel bom base nos total_late_orders e nomes de cada fornecedor
            # PedidosAtrasados.to_excel(f'total_late_orders{fornecedor[0]}.xlsx')

    def clickevent(self, click):
        global sendChoose
        sendChoose = click
        self.data_push()
        if(sendChoose=="corrective"):
            self.addCorrectiveWindow()
        elif(sendChoose=="preventive"):
            self.addPreventiveWindow()

    def addPreventiveWindow(self):
        self.pTopLevel = ""
        self.icon.iconbitmap("./fk-log.ico")
        self.pTopLevel.window.grab_set()
        self.pTopLevel.window.mainloop()
    
    def addCorrectiveWindow(self):
        #Window configuration
        self.cTopLevel = ctk.CTkToplevel()
        self.cTopLevel.title("Enviando emails corretivos...")
        self.cTopLevel.state('zoomed')
        self.cTopLevel.grab_set()
        self.cTopLevel.columnconfigure(0, weight=3)
        self.cTopLevel.columnconfigure(1, weight=3)
        self.cTopLevel.columnconfigure(2, weight=3)
        self.cTopLevel.rowconfigure(0, weight=5)
        self.cTopLevel.rowconfigure(1, weight=2)
        self.cTopLevel.rowconfigure(2, weight=3)
        self.cTopLevel.rowconfigure(3, weight=3)
        self.cTopLevel.rowconfigure(4, weight=3)
        self.cTopLevel.rowconfigure(5, weight=3)

        self.cListBox = CTkListbox(self.cTopLevel, width=700, height=250)
        for Name in correctiveSuppliersNames:
            self.cListBox.insert("END",Name)
        self.cListBox.grid(row=1, column=1, pady=10)

        self.suppliersNumbers = ctk.StringVar()
        self.suppliersNumbers.set(f"Total de Fornecedores: {self.cListBox.size()}")

        self.totalSuppliersLabel = ctk.CTkLabel(self.cTopLevel, textvariable=self.suppliersNumbers)
        self.totalSuppliersLabel.grid(row=0, column=1, pady=10, padx=10)

        self.restoreButton = ctk.CTkButton(self.cTopLevel, text="Restaurar", command=self.restoreListTopLevel)
        self.restoreButton.grid(row=2, column=0, sticky="W" ,padx=10)
        
        self.deleteButton = ctk.CTkButton(self.cTopLevel, text="Deletar", command=self.deleteSelectedItem, fg_color="#FF0000", text_color="white", hover_color="#990000")
        self.deleteButton.grid(row=2, column=2, pady=10, padx=10, sticky="E")

        self.cancelButton = ctk.CTkButton(self.cTopLevel,text="Cancelar", command=self.cTopLevel.destroy, width=300, height=50)
        self.cancelButton.grid(row=3, column=0, sticky="SW", padx=10)

        self.emailCcEntry = ctk.CTkEntry(self.cTopLevel, placeholder_text="Email:", width=200)
        self.emailCcEntry.grid(row=2, column=1)
        #testing bingind key presses
        #self.emailCcEntry.bind("<Return>", self.addCcEmail)

        self.emailCcListBox = CTkListbox(self.cTopLevel, width=300, height=200)
        self.emailCcListBox.grid(row=3, column=1)

        if(self.emailCcList==[]):
            pass
        else:
            for email in self.emailCcList:
                self.emailCcListBox.insert('end', email)

        self.emailCcAddButton = ctk.CTkButton(self.cTopLevel, text="Add Email +", command=self.addCcEmail)
        self.emailCcAddButton.grid(row=4, column=1)

        self.emailCcDeleteButton = ctk.CTkButton(self.cTopLevel, text="Deletar Email", command=self.deleteCcEmail, bg_color="RED")
        self.emailCcDeleteButton.grid(row=5, column=1, pady=10, padx=10)

        self.sendEmailsButton = ctk.CTkButton(self.cTopLevel, text="Enviar Email", command=lambda: self.sendCorrectiveEmail(correctiveSuppliersList), width=300, height=50)
        self.sendEmailsButton.grid(row=3, column=2, sticky="SE", padx=10)
    
    def selectedArchive(self, path):
        self.archiveTopLevel = ctk.CTkToplevel()
        self.archiveTopLevel.title("Arquivo selecionado")
        self.archiveTopLevel.geometry("300x200")
        self.archiveTopLevel.grab_set()

        self.playNotificationSound()

        #Shows "Selected Arqhive" in the window
        self.selectedArchiveLabel = ctk.CTkLabel(self.archiveTopLevel, text="Arquivo selecionado:", pady=10,padx=10)
        self.selectedArchiveLabel.pack(pady=10,padx=10)

        #splits the file path
        splitFilePath = path.split('/')
        splitLen = len(splitFilePath)-1
        fileName=splitFilePath[splitLen]
        
        #underline text configuration
        underlineText = ctk.CTkFont(underline=True)

        #Archive name label show
        self.selectedArchiveNameLabel = ctk.CTkLabel(self.archiveTopLevel, font=underlineText, text=fileName)
        self.selectedArchiveNameLabel.pack(pady=10, padx=10)

        self.okButton = ctk.CTkButton(self.archiveTopLevel, text="OK", command=self.archiveTopLevel.destroy)
        self.okButton.pack(pady=20, padx=20)

    def deleteSelectedItem(self):
        self.index = self.cListBox.curselection()
        self.cListBox.delete(self.index)

        for supplier in correctiveSuppliersList:
            if(supplier.Name==correctiveSuppliersList[self.index].Name):
                self.cDeletedSuppliers.append(correctiveSuppliersList[self.index])
                break

        print("Fornecedor deletado")
        print(correctiveSuppliersList[self.index].Name)

        correctiveSuppliersList.pop(self.index)

        print("Lista após o delete")
        for supplier in correctiveSuppliersList:
            print(supplier.Name)
        print(f"Tamanho lista: {len(correctiveSuppliersList)}")
        self.suppliersNumbers.set(f"Total de Fornecedores: {self.cListBox.size()}")

    def restoreListTopLevel(self):
        lastDeletedSupplier = len(self.cDeletedSuppliers)
        if(lastDeletedSupplier==0):
            self.emptyListTopLevel = ctk.CTkToplevel()
            self.emptyListTopLevel.title("Erro")
            self.emptyListTopLevel.geometry("300x200")
            self.emptyListTopLevel.grab_set()
            self.emptyListLabel = ctk.CTkLabel(self.emptyListTopLevel, text="Nenhum fornecedor foi deletado anteriormente")
            self.emptyListLabel.pack(pady=10, padx=10)
            self.emptyListButton = ctk.CTkButton(self.emptyListTopLevel, text="OK", command=self.emptyListTopLevel.destroy)
            self.emptyListButton.pack(pady=10, padx=10)

        elif(lastDeletedSupplier>0):
            self.deletedListTopLevel = ctk.CTkToplevel()
            self.deletedListTopLevel.title("Index")
            self.deletedListTopLevel.geometry("300x300")
            self.deletedListTopLevel.grab_set()

            self.ctkIndexList = CTkListbox(self.deletedListTopLevel)
            for fornecedor in self.cDeletedSuppliers:
                self.ctkIndexList.insert("END",f"{fornecedor.Name}")
            self.ctkIndexList.pack()

            self.button = ctk.CTkButton(self.deletedListTopLevel, width=100, height=100, text="OK", command=self.restoreListCommand)
            self.button.pack(pady=10) #to aqui

    def restoreListCommand(self):
            self.index = self.ctkIndexList.curselection()
            self.ctkIndexList.delete(self.index)

            self.cListBox.insert("END", self.cDeletedSuppliers[self.index].Name)

            for supplier in self.cDeletedSuppliers:
                if(supplier.Name == self.cDeletedSuppliers[self.index].Name):
                    correctiveSuppliersList.append(self.cDeletedSuppliers[self.index])
                    print(f"Fornecedor restaurado: {supplier.Name}")
                    print(f"Indice fornecedor: {self.index}")
                    self.cDeletedSuppliers.pop(self.index)
                    print(F"Tamanho da lista dos deletados após restaurar: {len(self.cDeletedSuppliers)}")
                    break
            if(self.cDeletedSuppliers==[]):
                print("Lista fornecedores após restaurar")
                for supplier in correctiveSuppliersList:
                    print(supplier.Name)
                
                self.deletedListTopLevel.destroy()

            self.suppliersNumbers.set(f"Total de Fornecedores: {self.cListBox.size()}")

    def addCcEmail(self):
        email = self.emailCcEntry.get()
        if(email!=""):
            self.emailCcList.append(email)
            self.emailCcListBox.insert("END", email)
            self.emailCcEntry.delete(0, 'end')
        else:
            print("Nada foi preenchido")

    def deleteCcEmail(self):
        self.index = self.emailCcListBox.curselection()
        self.emailCcListBox.delete(self.index)

        print("Deletando email:")
        for email in self.emailCcList:
            if(email==self.emailCcList[self.index]):
                print(f"Email deletado: {self.emailCcList[self.index]}")
                self.emailCcList.pop(self.index)
                break

    def playNotificationSound(self):
        pygame.mixer.music.load('./Notify.wav')
        pygame.mixer.music.play(loops=0)

    def sendCorrectiveEmail(self, suppliersList):
        outlook = win32.Dispatch("Outlook.Application")
        
        #ccEmail = ["glaucio.costa@fkgroup.com.br","luciana.santos@fkgroup.com.br", "guilherme.silva@fkgroup.com.br"]
        time.sleep(3)
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
                <h2>Favor confirmar a nova data de entrega desses pedidos que constam em atraso em nosso sistema: </h2>
                {lateOrdersHTML}

                <h3>Caso o pedido já tenha sido faturado ou despachado favor nos informar</h3>
            </body>
            </html>
            """
            email = outlook.CreateItem(0)
            time.sleep(1)
            #email.To = f'{supplier.Email}'
            email.To = "rafaelzinhobr159@gmail.com"

            if(self.emailCcList==[]):
                pass
            else:
                self.joinedEmail = "; ".join(self.emailCcList)
                email.Cc = self.joinedEmail

            email.Subject = f"Pedidos atrasados {supplier.Name}"
            email.HTMLBody = (correctiveEmailBody)
            time.sleep(1)
            email.send()
            time.sleep(2)
            print(f"Email enviado para: {supplier.Name}")

    def sendPreventiveEmail(self, suppliersList):
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

root = ctk.CTk()
root.iconbitmap(iconpath)
userinterface = interface(root)
root.mainloop()