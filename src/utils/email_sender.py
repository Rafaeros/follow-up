"""
    Module to send email to suppliers using Outlook COM
"""

import os
import time
import json
import win32com.client as win32
import pandas as pd
from src.core.config import ConfigManager

style: str = """
<style>
/* Aplica padding e cor de texto em todos os elementos */
* {
    padding: 5px;
    color: black;
    box-sizing: border-box; /* Inclui padding e border na largura e altura total do elemento */
}

/* Estilo para o cabeçalho da tabela */
thead {
    text-align: center;
    background-color: #5F9EA0; /* Cadetblue */
    color: white; /* Texto branco para contraste */
}

/* Estilo para linhas, cabeçalhos e células da tabela */
tr, th, td {
    text-align: center;
    vertical-align: middle; /* Alinhamento vertical ao centro */
    padding: 10px; /* Padding para melhorar espaçamento */
}

/* Estilo alternado para linhas da tabela */
tr:nth-child(even) {
    background-color: #f2f2f2; /* Fundo cinza claro */
}

/* Estilo para células específicas */
td:nth-child(5) {
    text-align: left;
    background-color: #FFCDD2; /* Vermelho claro para melhor legibilidade */
}

/* Estilo para a borda da tabela */
table {
    border-collapse: collapse; /* Remove espaçamento entre células */
    width: 100%; /* Largura total */
}

/* Estilo para bordas das células */
th, td {
    border: 1px solid #ddd; /* Borda cinza clara */
}

/* Estilo para hover em linhas da tabela */
tr:hover {
    background-color: #ddd; /* Fundo cinza claro ao passar o mouse */
}
</style>
"""


class EmailSender:
    def __init__(self, override_to: str | None = None, disable_cc: bool = False):
        self.config = ConfigManager()
        # For Outlook COM, we don't need SMTP credentials, but keep for compatibility
        self.override_to = override_to
        self.disable_cc = disable_cc
        self.emails_cc_file = os.path.join(
            os.path.dirname(__file__), "..", "..", "..", "emails_cc.json"
        )
        self.emails_cc = self.load_emails_cc()

    def load_emails_cc(self):
        if os.path.exists(self.emails_cc_file):
            try:
                with open(self.emails_cc_file, "r") as f:
                    return json.load(f)
            except json.JSONDecodeError:
                return []
        return []

    def send_corrective_email(self, suppliers: dict) -> None:
        """
        Send Corrective Email using Outlook COM
        """
        outlook = win32.Dispatch("Outlook.Application")
        error_log = pd.DataFrame(columns=['Name', 'Email', 'Error'])
        emails_cc_str = "; ".join(self.emails_cc) if self.emails_cc and not self.disable_cc else ""

        # Check if the folder exists, if not, create it
        folder_path: str = 'tmp'
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        time.sleep(3)
        for _, supplier_data in suppliers.items():
            if supplier_data["late_orders"] == []:
                continue

            late_orders_df = pd.DataFrame(
                supplier_data["late_orders"]).reset_index(drop=True)
            late_orders_df.index += 1
            late_orders_html = late_orders_df.to_html(
                col_space=50, justify='center')

            email_body: str = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    {style}
                </head>
                <body>
                    <p>Prezados</p>,

                    <p>
                        Gostaria de confirmar e validar a entrega dos materiais solicitados conforme o pedido enviado anteriormente.
                        Onde constam em atraso em nosso sistema, caso o pedido tenha sido faturado ou despachado favor nos informar.
                        Caso haja necessidade de ajustes ou informações adicionais, por favor, entrem em contato diretamente por este e-mail.
                    </p>

                    <p style="color: red">
                    Importante: Para garantir o cumprimento do cronograma, solicito que os itens sejam faturados com antecedência adequada,
                    permitindo que cheguem à nossa empresa na data prevista no pedido. Este procedimento é essencial para que possamos
                    manter o cronograma conforme o planejado.
                    </p>

                    <p>Caso haja necessidade de ajustes ou informações adicionais, por favor, entrem em contato diretamente por este e-mail.</p>

                    <p>Agradeço pela atenção e colaboração.</p>

                    <p>Atenciosamente, </p>

                    <h3>Pedidos: </h3>
                    {late_orders_html}
                </body>
                </html>
                """

            try:
                email = outlook.CreateItem(0)
                time.sleep(1)
                to_address = self.override_to or supplier_data["email"]
                email.To = to_address
                if emails_cc_str:
                    email.Cc = emails_cc_str
                email.Subject = f"Pedidos Atrasados {supplier_data['name']}"
                email.HTMLBody = email_body
                time.sleep(1)
                email.Send()
                time.sleep(2)

            except Exception as e:
                error_entry = pd.DataFrame({
                    'Name': [supplier_data['name']],
                    'Email': [supplier_data['email']],
                    'Error': [str(e)]
                })
                error_log = pd.concat([error_log, error_entry], ignore_index=True)
        if not error_log.empty:
            log_file_path = os.path.join(folder_path, 'error_corrective_log.xlsx')
            error_log.to_excel(log_file_path, index=False)
        outlook.Quit()

    def send_preventive_email(self, suppliers: dict) -> None:
        """
        Send Preventive Email using Outlook COM
        """
        outlook = win32.Dispatch("Outlook.Application")
        error_log = pd.DataFrame(columns=['Name', 'Email', 'Error'])
        emails_cc_str = "; ".join(self.emails_cc) if self.emails_cc and not self.disable_cc else ""

        folder_path: str = 'tmp'
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        time.sleep(3)
        for _, supplier_data in suppliers.items():
            if not supplier_data["preventive_orders"]:
                continue

            preventive_orders_df = pd.DataFrame(
                supplier_data["preventive_orders"]).reset_index(drop=True)
            preventive_orders_df.index += 1
            preventive_orders_html = preventive_orders_df.to_html(
                col_space=50, justify='center')

            email_body: str = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    {style}
                </head>
                <body>
                    <p>Prezados</p>,

                    <p>
                        Gostaria de confirmar e validar a entrega dos materiais solicitados conforme o pedido enviado anteriormente.
                        Este contato visa assegurar que todos os itens serão entregues conforme as datas estipuladas.
                        Conforme acordado, as entregas estão programadas para ocorrer dentro dos prazos indicados na tabela abaixo.
                        Solicito, por gentileza, a confirmação de que essas previsões estão de acordo com as expectativas e necessidades de sua equipe.
                    </p>

                    <p style="color: red">
                    Importante: Para garantir o cumprimento do cronograma, solicito que os itens sejam faturados com antecedência adequada,
                    permitindo que cheguem à nossa empresa na data prevista no pedido. Este procedimento é essencial para que possamos
                    manter o cronograma conforme o planejado.
                    </p>

                    <p>Caso haja necessidade de ajustes ou informações adicionais, por favor, entrem em contato diretamente por este e-mail.</p>

                    <p>Agradeço pela atenção e colaboração.</p>

                    <p>Atenciosamente, </p>

                    <h3>Pedidos: </h3>
                    {preventive_orders_html}
                </body>
                </html>
                """

            try:
                email = outlook.CreateItem(0)
                time.sleep(1)
                to_address = self.override_to or supplier_data["email"]
                email.To = to_address
                if emails_cc_str:
                    email.Cc = emails_cc_str
                email.Subject = f"Confirmação de Pedidos {supplier_data['name']}"
                email.HTMLBody = email_body
                time.sleep(1)
                email.Send()
                time.sleep(2)

            except Exception as e:
                error_entry = pd.DataFrame({
                    'Name': [supplier_data['name']],
                    'Email': [supplier_data['email']],
                    'Error': [str(e)]
                })
                error_log = pd.concat([error_log, error_entry], ignore_index=True)
        if not error_log.empty:
            log_file_path = os.path.join(folder_path, 'error_preventive_log.xlsx')
            error_log.to_excel(log_file_path, index=False)
        outlook.Quit()
