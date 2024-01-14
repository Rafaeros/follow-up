import pandas as pd
from Supplier import Supplier
import datetime as dt
from datetime import timedelta
import time
import win32com.client as win32
import customtkinter as ctk
from CTkListbox import *
from CTkMessagebox import CTkMessagebox
from playsound import playsound
from PIL import Image

# Getting today date
today_date = dt.datetime.now()
iconpath = "src/fk-logo.ico"

# Email style
style = """
<style>
* {
padding: 5px;
text-color: black;
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

class interface():
    def __init__(self, master):
        self.master = master
        #Declarating variables
        self.cDeletedSuppliers = []
        self.pDeletedSuppliers = [] 
        self.emailCcList = []
        self.leftCollumnsError = []
        self.dataError = ["Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam", "Nacionalidade", "Rateio", "Situação"]
        self.emailDataError = ["Nome", "Email"]
        self.listBoxTextColor = "black"
        self.isPreventiveEmailSended = False
        self.isCorrectiveEmailSended = False
        self.index = -5

        self.master.title("Follow Up F&K Group")
        self.master.protocol("WM_DELETE_WINDOW", self.onClosing)
        self.master.geometry("500x500")

        self.step_1 = ctk.CTkLabel(self.master, text="1° Passo")
        self.step_1.pack(pady=30)

        self.emailDialogButton = ctk.CTkButton(
            self.master, text="Adicionar Arquivo C/ Emails", command=self.add_email_file)
        self.emailDialogButton.pack(pady=10)

        self.step_2 = ctk.CTkLabel(self.master, text="2° Passo")
        self.step_2.pack(pady=10)

        self.fileDialogButton = ctk.CTkButton(
            self.master, text="Adicionar Arquivo C/ Pedidos", command=self.add_file)
        self.fileDialogButton.pack(pady=10)

        self.step_3 = ctk.CTkLabel(self.master, text="3° Passo")
        self.step_3.pack(pady=10)

        self.sendlateOrder_emails = ctk.CTkButton(self.master, text="Enviar Emails Atrasados", command=lambda m="corrective": self.clickevent(m))
        self.sendlateOrder_emails.pack(pady=10)

        self.sendPreventive_emails = ctk.CTkButton(
            self.master, text="Enviar Emails Preventivos", command=lambda m="preventive": self.clickevent(m))
        self.sendPreventive_emails.pack(pady=10)

        self.lightImage = ctk.CTkImage(Image.open("./src/light.png"), size=(100,50))
        self.darkImage = ctk.CTkImage(Image.open("./src/dark.png"), size=(100,50))

        self.toggleThemeButton = ctk.CTkButton(self.master, text="", image=self.lightImage, bg_color="#EBEBEB", fg_color="#EBEBEB", width=40, height=20, command=self.toggleTheme)
        self.toggleThemeButton['border']=0
        self.toggleThemeButton.pack(pady=20)

    def toggleTheme(self):
        currentColor = self.toggleThemeButton.cget("bg_color")
        if(currentColor=='#EBEBEB'):
            self.toggleThemeButton.configure(image=self.darkImage)
            self.toggleThemeButton.configure(bg_color='#242424')
            self.toggleThemeButton.configure(fg_color='#242424')
            self.listBoxTextColor = "white"
            self.appearance = ctk.set_appearance_mode("Dark")
        else:
            self.toggleThemeButton.configure(image=self.lightImage)
            self.toggleThemeButton.configure(bg_color='#EBEBEB')
            self.toggleThemeButton.configure(fg_color='#EBEBEB')
            self.listBoxTextColor = "black"
            self.appearance = ctk.set_appearance_mode("Light")

    def add_email_file(self):
        global email_data_filepath
        email_data_filepath = ctk.filedialog.askopenfilename()
        email_data_filepath = "".join(email_data_filepath)
        if(email_data_filepath!=""):
            self.selectedArchive(email_data_filepath)
        else:
            self.EmptyFilePathPopUp()

    def add_file(self):
        global orders_data_filepath
        orders_data_filepath = ctk.filedialog.askopenfilename()
        orders_data_filepath = "".join(orders_data_filepath)
        if(orders_data_filepath!=""):
            self.selectedArchive(orders_data_filepath)
        else:
            self.EmptyFilePathPopUp()

    def format_data(self, Orders):
        Orders.pop(Orders.columns[0])

        Orders.index += 1

        Orders['Data de entrega'] = pd.to_datetime(
            Orders['Data de entrega'], format='%d/%m/%Y')

        Orders['Data de entrega'] = Orders["Data de entrega"].dt.strftime("%d/%m/%Y")

    def dataValidation(self, emailData, ordersData):
        for error in self.emailDataError:
            if error in emailData.columns:
                pass
            else:
                pass
                self.leftCollumnsError.append(error)
        for error in self.dataError:
            if error in ordersData.columns:
                pass
            else:
                self.leftCollumnsError.append(error)

        if(self.leftCollumnsError==[]):
            pass
        else:
            self.dataWarnTopLevel(self.leftCollumnsError)

    def dataWarnTopLevel(self, errorsList):
        self.errorWarnToplevel = ctk.CTkToplevel()
        self.errorWarnToplevel.title("Erro: planilha sem as colunas necessárias")
        self.errorWarnToplevel.geometry("400x200")
        self.errorWarnToplevel.grab_set()
        errorsWarnText = ", ".join(errorsList)
        self.errorsWarnLabel = ctk.CTkLabel(self.errorWarnToplevel, text="Colunas não encontradas: ", wraplength=400)
        self.errorsWarnLabel.pack(pady=10)
        self.errorsListLabel = ctk.CTkLabel(self.errorWarnToplevel, text=f"{errorsWarnText}", wraplength=300)
        self.errorsListLabel.pack(pady=10)
        self.closeErrorTopLevelButton = ctk.CTkButton(self.errorWarnToplevel, text="OK", command=self.errorWarnToplevel.destroy)
        self.closeErrorTopLevelButton.pack(pady=10)
        self.errorWarnToplevel.wait_window()
        self.leftCollumnsError.clear()

    def data_push(self):
        suppliersData = pd.read_excel(email_data_filepath)
        totalOrders = pd.read_excel(orders_data_filepath)

        self.dataValidation(suppliersData, totalOrders)

        emails_data = suppliersData[["Nome", "Email"]]

        totalOrders = totalOrders[totalOrders['Situação'] != 'Envio pendente']
        totalOrders = totalOrders[totalOrders['Nacionalidade'] == 'Brasil']
        MP_filter = ['MATERIA-PRIMA', 'MATERIA PRIMA INDUSTRIALIZAÇÃO',
                    'MATERIAL DE USO E CONSUMO']
        totalOrders = totalOrders[totalOrders['Rateio'].isin(MP_filter)]

        totalOrders['Data de entrega'] = pd.to_datetime(
            totalOrders['Data de entrega'], format='%d/%m/%Y')
        
        self.ordersReport = totalOrders

        # Late Orders for corrective treatment
        global lastDay
        lastDay = today_date - timedelta(days=1)
        total_late_orders = totalOrders[totalOrders['Data de entrega'] < lastDay]

        # Ten days ahead Orders for preventive preventive treatment
        global dateTenDaysAhead
        dateTenDaysAhead = today_date + timedelta(days=11)
        dateMask = (totalOrders['Data de entrega'] > today_date) & (totalOrders['Data de entrega'] <= dateTenDaysAhead)
        orders_tenDaysAhead = totalOrders.loc[dateMask]

        if (sendChoose == "corrective"):
            global correctiveSuppliersNamesList
            global correctiveSuppliersNames
            global correctiveSuppliersList

            correctiveSuppliersNames = []

            correctiveSuppliersNamesList = total_late_orders.loc[:, ['Fornecedor']].drop_duplicates(
                subset="Fornecedor", keep="first").values.tolist()
            
            for Name in correctiveSuppliersNamesList:
                correctiveSuppliersNames.append(Name[0])

            correctiveSuppliersList = []
            for Name in correctiveSuppliersNames:
                lateOrders = total_late_orders.loc[total_late_orders['Fornecedor'] == Name, [
                    "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()
                self.format_data(lateOrders)

                cCurrent_email = emails_data.loc[emails_data['Nome'] == Name, [
                    "Email"]].to_string(index=False, header=False)
                            
                splitcCurrent_email = cCurrent_email.split(sep=",")

                joincCurrent_email = "; ".join(splitcCurrent_email)

                correctiveSuppliersList.append(
                    Supplier(Name, f"{joincCurrent_email}", lateOrders))

        elif (sendChoose == "preventive"):
            global preventiveSuppliersNamesList
            global preventiveSuppliersNames
            global preventiveSuppliersList

            preventiveSuppliersNames = []

            preventiveSuppliersNamesList = orders_tenDaysAhead.loc[:, ['Fornecedor']].drop_duplicates(
                subset="Fornecedor", keep="first").values.tolist()
            
            for Name in preventiveSuppliersNamesList:
                preventiveSuppliersNames.append(Name[0])
            
            preventiveSuppliersList = []
            for Name in preventiveSuppliersNames:

                preventiveOrders = orders_tenDaysAhead.loc[orders_tenDaysAhead['Fornecedor'] == Name, [
                    "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]].reset_index()

                self.format_data(preventiveOrders)

                pCurrent_email = emails_data.loc[emails_data['Nome'] == Name , ["Email"]]
                pCurrent_email = pCurrent_email.to_string(header=False, index=False)

                preventiveSuppliersList.append(
                    Supplier(Name, pCurrent_email, preventiveOrders))

    def clickevent(self, click):
        global sendChoose
        sendChoose = click
        self.data_push()
        if(sendChoose=="corrective"):
            self.addCorrectiveWindow()
        elif(sendChoose=="preventive"):
            self.addPreventiveWindow()

    def addPreventiveWindow(self):
        self.pTopLevel = ctk.CTkToplevel()
        self.pTopLevel.title("Enviando emails preventivos...")
        self.pTopLevel.state('zoomed')
        self.pTopLevel.grab_set()
        self.pTopLevel.columnconfigure(0, weight=3)
        self.pTopLevel.columnconfigure(1, weight=3)
        self.pTopLevel.columnconfigure(2, weight=3)
        self.pTopLevel.rowconfigure(0, weight=5)
        self.pTopLevel.rowconfigure(1, weight=2)
        self.pTopLevel.rowconfigure(2, weight=3)
        self.pTopLevel.rowconfigure(3, weight=3)
        self.pTopLevel.rowconfigure(4, weight=3)
        self.pTopLevel.rowconfigure(5, weight=3)

        self.pListBox = CTkListbox(self.pTopLevel, width=700, height=250, text_color=f"{self.listBoxTextColor}")
        for Name in preventiveSuppliersNames:
            self.pListBox.insert("END",Name)
        self.pListBox.grid(row=1, column=1, pady=10)

        self.pSuppliersNumbers = ctk.StringVar()
        self.pSuppliersNumbers.set(f"Total de Fornecedores: {self.pListBox.size()}")

        self.pTotalSuppliersLabel = ctk.CTkLabel(self.pTopLevel, textvariable=self.pSuppliersNumbers)
        self.pTotalSuppliersLabel.grid(row=0, column=1, pady=10, padx=10)

        self.pRestoreButton = ctk.CTkButton(self.pTopLevel, text="Restaurar", command=self.restorePreventiveListTopLevel)
        self.pRestoreButton.grid(row=2, column=0, sticky="W" ,padx=10)
        
        self.pDeleteButton = ctk.CTkButton(self.pTopLevel, text="Deletar", command=self.deletePreventiveSelectedItem, fg_color="#FF0000", text_color="white", hover_color="#990000")
        self.pDeleteButton.grid(row=2, column=2, pady=10, padx=10, sticky="E")

        self.pCancelButton = ctk.CTkButton(self.pTopLevel,text="Cancelar", command=self.pTopLevel.destroy, width=300, height=50)
        self.pCancelButton.grid(row=3, column=0, sticky="SW", padx=10)

        self.emailCcEntry = ctk.CTkEntry(self.pTopLevel, placeholder_text="Digite o Email em cópia:", width=200)
        self.emailCcEntry.grid(row=2, column=1)

        #testing bingind key presses
        #self.emailCcEntry.bind("<Return>", self.addCcEmail)

        self.emailCcListBox = CTkListbox(self.pTopLevel, width=300, height=200, text_color=f"{self.listBoxTextColor}")
        self.emailCcListBox.grid(row=3, column=1)

        if(self.emailCcList==[]):
            pass
        else:
            for email in self.emailCcList:
                self.emailCcListBox.insert('end', email)

        self.emailCcAddButton = ctk.CTkButton(self.pTopLevel, text="Add Email +", command=self.addCcEmail)
        self.emailCcAddButton.grid(row=4, column=1)

        self.emailCcDeleteButton = ctk.CTkButton(self.pTopLevel, text="Deletar Email", command=self.deleteCcEmail, bg_color="RED")
        self.emailCcDeleteButton.grid(row=5, column=1, pady=10, padx=10)

        self.sendPEmailsButton = ctk.CTkButton(self.pTopLevel, text="Enviar Email", command=lambda: self.sendPreventiveEmail(preventiveSuppliersList), width=300, height=50)
        self.sendPEmailsButton.grid(row=3, column=2, sticky="SE", padx=10)

    def addCorrectiveWindow(self):
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

        self.cListBox = CTkListbox(self.cTopLevel, width=700, height=250, text_color=f"{self.listBoxTextColor}")
        for Name in correctiveSuppliersNames:
            self.cListBox.insert("END",Name)
        self.cListBox.grid(row=1, column=1, pady=10)

        self.cSuppliersNumbers = ctk.StringVar()
        self.cSuppliersNumbers.set(f"Total de Fornecedores: {self.cListBox.size()}")

        self.totalSuppliersLabel = ctk.CTkLabel(self.cTopLevel, textvariable=self.cSuppliersNumbers)
        self.totalSuppliersLabel.grid(row=0, column=1, pady=10, padx=10)

        self.restoreButton = ctk.CTkButton(self.cTopLevel, text="Restaurar", command=self.restoreCorrectiveListTopLevel)
        self.restoreButton.grid(row=2, column=0, sticky="W" ,padx=10)
        
        self.deleteButton = ctk.CTkButton(self.cTopLevel, text="Deletar", command=self.deleteCorrectiveSelectedItem, fg_color="#FF0000", text_color="white", hover_color="#990000")
        self.deleteButton.grid(row=2, column=2, pady=10, padx=10, sticky="E")

        self.cancelButton = ctk.CTkButton(self.cTopLevel,text="Cancelar", command=self.cTopLevel.destroy, width=300, height=50)
        self.cancelButton.grid(row=3, column=0, sticky="SW", padx=10)

        self.emailCcEntry = ctk.CTkEntry(self.cTopLevel, placeholder_text="Digite o email em cópia:", width=200)
        self.emailCcEntry.grid(row=2, column=1)
        #testing bingind key presses
        #self.emailCcEntry.bind("<Return>", self.addCcEmail)

        self.emailCcListBox = CTkListbox(self.cTopLevel, width=300, height=200, text_color=f"{self.listBoxTextColor}")
        self.emailCcListBox.grid(row=3, column=1)

        if(self.emailCcList==[]):
            pass
        else:
            for email in self.emailCcList:
                self.emailCcListBox.insert('end', email)

        self.emailCcAddButton = ctk.CTkButton(self.cTopLevel, text="Add Email +", command=self.addCcEmail)
        self.emailCcAddButton.grid(row=4, column=1)

        self.emailCcDeleteButton = ctk.CTkButton(self.cTopLevel, text="Deletar Email", command=self.deleteCcEmail)
        self.emailCcDeleteButton.grid(row=5, column=1, pady=10, padx=10)

        self.sendCEmailsButton = ctk.CTkButton(self.cTopLevel, text="Enviar Email", command=lambda: self.sendCorrectiveEmail(correctiveSuppliersList), width=300, height=50)
        self.sendCEmailsButton.grid(row=3, column=2, sticky="SE", padx=10)
    
    def selectedArchive(self, path):
        self.archiveTopLevel = ctk.CTkToplevel()
        self.archiveTopLevel.title("Arquivo selecionado")
        self.archiveTopLevel.geometry("300x200")
        self.archiveTopLevel.grab_set()

        playsound("./src/Notify.wav", block=False)

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

    def deletePreventiveSelectedItem(self):
        self.index = self.pListBox.curselection()
        self.pListBox.delete(self.index)

        for supplier in preventiveSuppliersList:
            if(supplier.Name==preventiveSuppliersList[self.index].Name):
                self.pDeletedSuppliers.append(preventiveSuppliersList[self.index])
                break

        preventiveSuppliersList.pop(self.index)

        self.pSuppliersNumbers.set(f"Total de Fornecedores: {self.pListBox.size()}")

    def deleteCorrectiveSelectedItem(self):
        self.index = self.cListBox.curselection()
        self.cListBox.delete(self.index)

        for supplier in correctiveSuppliersList:
            if(supplier.Name==correctiveSuppliersList[self.index].Name):
                self.cDeletedSuppliers.append(correctiveSuppliersList[self.index])
                break

        correctiveSuppliersList.pop(self.index)

        self.cSuppliersNumbers.set(f"Total de Fornecedores: {self.cListBox.size()}")

    def restoreCorrectiveListTopLevel(self):
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
            self.deletedListTopLevel.title("Fornecedores deletados")
            self.deletedListTopLevel.geometry("400x400")
            self.deletedListTopLevel.grab_set()

            self.correctiveDeletedSupplierListBox = CTkListbox(self.deletedListTopLevel, text_color=f"{self.listBoxTextColor}")
            for fornecedor in self.cDeletedSuppliers:
                self.correctiveDeletedSupplierListBox.insert("END",f"{fornecedor.Name}")
            self.correctiveDeletedSupplierListBox.pack()

            self.button = ctk.CTkButton(self.deletedListTopLevel, width=100, height=100, text="OK", command=self.restoreCorrectiveListCommand)
            self.button.pack(pady=10) #to aqui

    def restoreCorrectiveListCommand(self):
            self.index = self.correctiveDeletedSupplierListBox.curselection()
            self.correctiveDeletedSupplierListBox.delete(self.index)

            self.cListBox.insert("END", self.cDeletedSuppliers[self.index].Name)

            for supplier in self.cDeletedSuppliers:
                if(supplier.Name == self.cDeletedSuppliers[self.index].Name):
                    correctiveSuppliersList.append(self.cDeletedSuppliers[self.index])
                    self.cDeletedSuppliers.pop(self.index)
                    break
            if(self.cDeletedSuppliers==[]):
                self.deletedListTopLevel.destroy()

            self.cSuppliersNumbers.set(f"Total de Fornecedores: {self.cListBox.size()}")

    def restorePreventiveListTopLevel(self):
        pLastDeletedSupplier = len(self.pDeletedSuppliers)
        if(pLastDeletedSupplier==0):
            self.emptyPListTopLevel = ctk.CTkToplevel()
            self.emptyPListTopLevel.title("Erro")
            self.emptyPListTopLevel.geometry("300x200")
            self.emptyPListTopLevel.grab_set()
            self.emptyPListLabel = ctk.CTkLabel(self.emptyPListTopLevel, text="Nenhum fornecedor foi deletado anteriormente")
            self.emptyPListLabel.pack(pady=10, padx=10)
            self.emptyPListButton = ctk.CTkButton(self.emptyPListTopLevel, text="OK", command=self.emptyPListTopLevel.destroy)
            self.emptyPListButton.pack(pady=10, padx=10)

        elif(pLastDeletedSupplier>0):
            self.deletedPListTopLevel = ctk.CTkToplevel()
            self.deletedPListTopLevel.title("Fornecedores deletados")
            self.deletedPListTopLevel.geometry("400x400")
            self.deletedPListTopLevel.grab_set()

            self.preventiveDeletedSupplierListBox = CTkListbox(self.deletedPListTopLevel, text_color=f"{self.listBoxTextColor}")
            for supplier in self.pDeletedSuppliers:
                self.preventiveDeletedSupplierListBox.insert("END",f"{supplier.Name}")
            self.preventiveDeletedSupplierListBox.pack()
            self.pButton = ctk.CTkButton(self.deletedPListTopLevel, width=100, height=100, text="OK", command=self.restorePreventiveListCommand)
            self.pButton.pack(pady=10)

    def restorePreventiveListCommand(self):
            self.index = self.preventiveDeletedSupplierListBox.curselection()
            self.preventiveDeletedSupplierListBox.delete(self.index)

            self.pListBox.insert("END", self.pDeletedSuppliers[self.index].Name)

            for supplier in self.pDeletedSuppliers:
                if(supplier.Name == self.pDeletedSuppliers[self.index].Name):
                    preventiveSuppliersList.append(self.pDeletedSuppliers[self.index])
                    self.pDeletedSuppliers.pop(self.index)
                    break
            if(self.pDeletedSuppliers==[]):
                self.deletedPListTopLevel.destroy()

            self.pSuppliersNumbers.set(f"Total de Fornecedores: {self.pListBox.size()}")

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

        for email in self.emailCcList:
            if(email==self.emailCcList[self.index]):
                self.emailCcList.pop(self.index)
                break

    def sendCorrectiveEmail(self, suppliersList):
        outlook = win32.Dispatch("Outlook.Application")
        
        time.sleep(3)
        for supplier in suppliersList[:]:
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
            email.To = f'{supplier.Email}'

            if(self.emailCcList==[]):
                pass
            else:
                self.joinedEmail = "; ".join(self.emailCcList)
                email.Cc = self.joinedEmail

            email.Subject = f"Pedidos atrasados {supplier.Name}"
            email.HTMLBody = (correctiveEmailBody)
            time.sleep(1)
            email.Send()
            time.sleep(2)

            suppliersList.pop(0)
            self.cListBox.delete(0)
            self.cSuppliersNumbers.set(f"Total de Fornecedores: {self.cListBox.size()}")

        if(suppliersList==[]):
            playsound("./src/Notify.wav", block=False)
            self.isCorrectiveEmailSended = True
            self.EmailsSendPopUp()
    
    def EmailsSendPopUp(self):
        emailSendMessage = CTkMessagebox(title="Concluído!", message="Todos os emails foram enviados com sucesso!", option_1="Ok", icon="check", text_color=f"{self.listBoxTextColor}")

    def EmptyFilePathPopUp(self):
        self.emptyFilePathTopLevel = ctk.CTkToplevel()
        self.emptyFilePathTopLevel.title("Atenção")
        self.emptyFilePathTopLevel.geometry("300x150")
        self.emptyFilePathTopLevel.grab_set()

        self.emptyFilePathWarn = ctk.CTkLabel(self.emptyFilePathTopLevel, text="Nenhum arquivo foi selecionado!")
        self.emptyFilePathWarn.pack(pady=10)

        self.emptyFilePathButton = ctk.CTkButton(self.emptyFilePathTopLevel, text="OK", command=self.emptyFilePathTopLevel.destroy)
        self.emptyFilePathButton.pack(pady=10)

    def sendPreventiveEmail(self, suppliersList):
        outlook = win32.Dispatch("Outlook.Application")

        time.sleep(3)
        for supplier in suppliersList[:]:
            pLateOrdersHTML = supplier.TotalOrders.to_html(
                col_space=50, justify='center')
            preventiveEmailBody = f"""
            <!DOCTYPE html>
            <html>
            <head>
                {style}
            </head>
            <body>
                <h1>Olá:{supplier.Name}</h1>
                <h2>Favor validar e confirmar a entrega desses pedidos conforme as datas previstas para estar aqui na F&K Group: </h2>
                {pLateOrdersHTML}
            </body>
            </html>
            """
            email = outlook.CreateItem(0)
            time.sleep(1)
            email.To = f'{supplier.Email}'

            if(self.emailCcList==[]):
                pass
            else:
                self.joinedEmail = "; ".join(self.emailCcList)
                email.Cc = self.joinedEmail

            email.Subject = f"Entrega Pedidos: {supplier.Name}"
            email.HTMLBody = (preventiveEmailBody)
            time.sleep(1)
            email.Send()
            time.sleep(2)

            suppliersList.pop(0)
            self.pListBox.delete(0)
            self.pSuppliersNumbers.set(f"Total de Fornecedores: {self.pListBox.size()}")

        if(suppliersList==[]):
            playsound("./src/Notify.wav", block=False)
            self.isPreventiveEmailSended = True
            self.EmailsSendPopUp()

    def formatReportDate(self, ordersReport):
        ordersReport['Data de entrega'] = pd.to_datetime(
            ordersReport['Data de entrega'], format='%d/%m/%Y')
        ordersReport['Data de entrega'] = ordersReport["Data de entrega"].dt.strftime("%d/%m/%Y")

    def EmailSendReport(self):
        formatDate = today_date.strftime("%d-%m-%Y")
        
        if(self.isPreventiveEmailSended==True and self.isCorrectiveEmailSended==True):

            reportDateMask = (self.ordersReport['Data de entrega'] < lastDay) | (self.ordersReport['Data de entrega'] <= dateTenDaysAhead)
            self.ordersReport = self.ordersReport.loc[reportDateMask]
            self.formatReportDate(self.ordersReport)
            self.ordersReport.to_excel(f"EmailsEnviados(Corretivo-Preventivo) {formatDate}.xlsx", index=False, sheet_name=f"Relatório {formatDate}")

        elif(self.isCorrectiveEmailSended==True):

            self.ordersReport = self.ordersReport[self.ordersReport["Data de entrega"] < lastDay]
            self.formatReportDate(self.ordersReport)
            self.ordersReport.to_excel(f"EmailsEnviados(Corretivo) {formatDate}.xlsx", index=False, sheet_name=f"Relatório {formatDate}")

        elif(self.isPreventiveEmailSended==True):

            preventiveDateMask = (self.ordersReport['Data de entrega'] > today_date) & (self.ordersReport['Data de entrega'] <= dateTenDaysAhead)
            self.ordersReport = self.ordersReport.loc[preventiveDateMask]
            self.formatReportDate(self.ordersReport)
            self.ordersReport.to_excel(f"EmailsEnviados(Preventivo) {formatDate}.xlsx", index=False, sheet_name=f"Relatório {formatDate}")

        else:
            noSendedEmailWarn = CTkMessagebox(title="Atenção", text_color=f"{self.listBoxTextColor}", message="Nenhum email enviado, encerrando o programa!", icon="info", option_1="Ok")

    def onClosing(self):
        closeMessage = CTkMessagebox(text_color=f"{self.listBoxTextColor}", title="Fechar?", message="Tem certeza que deseja encerrar o programa?", icon="question", option_1="Cancelar", option_2="Fechar")
        response = closeMessage.get()
        if(response=="Fechar"):
            self.EmailSendReport()
            root.destroy()
        
root = ctk.CTk()
root.iconbitmap(iconpath)
userinterface = interface(root)
root.mainloop()