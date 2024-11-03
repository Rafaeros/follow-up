"""
Main application interface.
"""

import customtkinter as ctk
from CTkMessagebox import CTkMessagebox as CTkmsg
from CTkListbox import CTkListbox as CTklist
from supplier_data import Suppliers
from send_email import send_corrective_email, send_preventive_email


class Interface(ctk.CTk):
    """
    Main application interface.
    """
    font: tuple

    def __init__(self) -> None:
        """
        Initializes the main application window.

        :param master: The parent window.
        :type master: ctk.CTk
        """
        super().__init__()
        self.title("Pedidos Atrasados")
        self.geometry("1000x800")
        self.index: int = 0
        self.orders_file_path: str = ""
        self.orders_file_name = ctk.StringVar()
        self.orders_file_name.set("Arquivo Selecionado: ")
        self.emails_file_path: str = ""
        self.email_file_name = ctk.StringVar()
        self.email_file_name.set("Arquivo Selecionado: ")
        self.emails_cc: list[str] = []
        self.corrective_check_var = ctk.StringVar()
        self.corrective_check_var.set("on")
        self.preventive_check_var = ctk.StringVar()
        self.preventive_check_var.set("on")
        self.font: tuple = ("Arial", 25)

        # Orders File layout
        self.order_label = ctk.CTkLabel(
            self, text="Adicionar arquivo de pedidos: ", font=self.font,)
        self.order_label.grid(row=1, column=1, padx=50, pady=10, sticky="w")
        self.add_orders_file_button = ctk.CTkButton(
            self, text="Adicionar", font=self.font, fg_color="green", width=200,
            command=self.add_orders_file)
        self.add_orders_file_button.grid(row=1, column=2, pady=10, sticky="w")
        self.selected_order_file_label = ctk.CTkLabel(
            self, textvariable=self.orders_file_name, font=self.font)
        self.selected_order_file_label.grid(
            row=2, column=1, padx=50, pady=10, sticky="w")

        # Emails File layout
        self.email_label = ctk.CTkLabel(
            self, text="Adicionar arquivo de emails: ", font=self.font)
        self.email_label.grid(row=3, column=1, padx=50, pady=10, sticky="w")
        self.add_email_file_button = ctk.CTkButton(
            self, text="Adicionar", font=self.font, fg_color="green", width=200,
            command=self.add_email_file)
        self.add_email_file_button.grid(row=3, column=2, pady=50, sticky="w")
        self.selected_email_file_label = ctk.CTkLabel(
            self, textvariable=self.email_file_name, font=self.font)
        self.selected_email_file_label.grid(
            row=4, column=1, padx=50, pady=10, sticky="w")

        # Checkboxes
        self.corrective_checkbox = ctk.CTkCheckBox(self, text="Corretivos",
            font=self.font, variable=self.corrective_check_var,
                onvalue='on', offvalue='off')
        self.corrective_checkbox.grid(
            row=5, column=1, padx=50, pady=10, sticky="w")
        self.preventive_checkbox = ctk.CTkCheckBox(self, text="Preventivos",
            font=self.font, variable=self.preventive_check_var,
                onvalue='on', offvalue='off')

        self.preventive_checkbox.grid(
            row=5, column=2, padx=10, pady=10, sticky="w")

        # CC Emails Layout
        self.email_cc_label = ctk.CTkLabel(
            self, text="Emails em Cópia", font=self.font)
        self.email_cc_label.grid(row=6, column=1, padx=50, pady=10, sticky="w")
        self.email_cc_entry = ctk.CTkEntry(
            self, font=self.font, placeholder_text="Email em cópia")
        self.email_cc_entry.bind("<Return>", self.add_cc_emails)
        self.email_cc_entry.bind("<Delete>", self.remove_cc_emails)
        self.email_cc_entry.grid(
            row=7, column=1, padx=50, pady=10, sticky="ew")
        self.email_cc_list = CTklist(
            self, width=400, height=200, font=self.font)
        self.email_cc_list.grid(row=8, column=1, padx=50,
                                pady=10, sticky="ew", rowspan=2)

        self.add_cc_emails_button = ctk.CTkButton(
            self, text="Adicionar Email", font=self.font, fg_color="green",
            command=self.add_cc_emails)
        self.add_cc_emails_button.grid(
            row=8, column=2, padx=10, pady=10, sticky="ew")

        self.remove_cc_emails_button = ctk.CTkButton(
            self, text="Remover Email", font=self.font, fg_color="red",
            command=self.remove_cc_emails)
        self.remove_cc_emails_button.grid(
            row=9, column=2, padx=10, pady=10, sticky="ew")

        # Send emails button
        self.send_emails_button = ctk.CTkButton(
            self, text="Enviar Emails", font=self.font, command=self.send_email, width=300)
        self.send_emails_button.grid(
            row=10, column=1, columnspan=2, padx=50, pady=10, sticky="ew")

    def add_cc_emails(self, event=None) -> None:
        """
        Add Cc Emails in a list
        """
        print(event)
        email = self.email_cc_entry.get()
        if email == "":
            CTkmsg(self, title="Erro", message="Email em branco", icon="warning")
            return

        self.email_cc_list.insert("END", email)
        self.emails_cc.append(email)
        self.email_cc_entry.delete(0, 'end')

    def remove_cc_emails(self, event=None) -> None:
        """
        Remove Cc Emails from a list
        """
        print(event)
        self.index = self.email_cc_list.curselection()
        self.email_cc_list.delete(self.index)
        self.emails_cc.pop(self.index)

    def add_orders_file(self) -> None:
        """
        Opens a file dialog to select the orders excel (*.xlsx) file.
        """
        self.orders_file_path = ctk.filedialog.askopenfilename(
            filetypes=[("Arquivos de excel", "*.xlsx")])

        self.orders_file_name.set(f"Arquivo Selecionado: {
                                  self.orders_file_path.split('/')[-1]}")

    def add_email_file(self) -> None:
        """
        Opens a file dialog to select the emails excel (*.xlsx) file.
        """
        self.emails_file_path = ctk.filedialog.askopenfilename(
            filetypes=[("Arquivos de excel", "*.xlsx")])

        self.email_file_name.set(f"Arquivo Selecionado: {
                                 self.emails_file_path.split('/')[-1]}")

    def send_email(self) -> None:
        """
        Gets the data from the selected files and sends emails.
        """
        if self.orders_file_path == "" or self.emails_file_path == "":
            CTkmsg(self, title="Erro",
                   message="Nenhum arquivo selecionado", icon="warning")
            return

        if (self.corrective_check_var.get() == 'off' and self.preventive_check_var.get() == 'off'):
            CTkmsg(self, title="Erro",
                   message="Nenhuma opção de envio de email selecionada", icon="warning")
            return

        s = Suppliers()
        s.get_data_from_file(self.orders_file_path, self.emails_file_path)
        suppliers: dict = s.to_json()

        if self.corrective_check_var.get() == 'on':
            send_corrective_email(suppliers, self.emails_cc)
        if self.preventive_check_var.get() == 'on':
            send_preventive_email(suppliers, self.emails_cc)

        CTkmsg(self, title="Sucesso",
               message="Emails enviados com sucesso!", icon="check")
