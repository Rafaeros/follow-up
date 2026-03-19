"""
Module to send email to suppliers using Microsoft Graph API
"""

import os
import json
import logging
import time
import pandas as pd
import requests
from src.core.config import ConfigManager
from src.utils.ms_graph_auth import MSGraphAuth

logger = logging.getLogger(__name__)

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
        self.override_to = override_to
        self.disable_cc = disable_cc
        self.emails_cc_file = os.path.join(
            os.path.dirname(__file__), "..", "..", "emails_cc.json"
        )
        self.emails_cc = self.load_emails_cc()

        # Initialize MS Graph Auth
        ms_config = self.config.get_ms_graph_config()
        self.auth = MSGraphAuth(
            client_id=ms_config.get("ms_graph_client_id"),
            tenant_id=ms_config.get("ms_graph_tenant_id"),
        )

    def load_emails_cc(self):
        if os.path.exists(self.emails_cc_file):
            try:
                with open(self.emails_cc_file, "r") as f:
                    return json.load(f)
            except json.JSONDecodeError:
                return []
        return []

    def _send_request(self, email_to, subject, body_html, cc_recipients=None):
        """
        Internal helper to send the email via MS Graph API
        """
        token = self.auth.get_access_token()
        url = "https://graph.microsoft.com/v1.0/me/sendMail"

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        to_recipients = [{"emailAddress": {"address": email_to}}]
        cc_list = []
        if cc_recipients and not self.disable_cc:
            cc_list = [{"emailAddress": {"address": cc}} for cc in cc_recipients]

        email_data = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body_html},
                "toRecipients": to_recipients,
                "ccRecipients": cc_list,
            },
            "saveToSentItems": True,
        }

        response = requests.post(url, headers=headers, json=email_data)
        if response.status_code == 202:
            logger.info(f"Email sent successfully to {email_to}")
        else:
            logger.error(
                f"Error sending email to {email_to}: {response.status_code} - {response.text}"
            )
            raise Exception(
                f"Graph API Error: {response.status_code} - {response.text}"
            )

    def send_corrective_email(self, suppliers: dict) -> None:
        """
        Send Corrective Email using Microsoft Graph API
        """
        error_log = pd.DataFrame(columns=["Name", "Email", "Error"])

        folder_path: str = "tmp"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        for _, supplier_data in suppliers.items():
            if supplier_data["late_orders"] == []:
                continue

            late_orders_df = pd.DataFrame(supplier_data["late_orders"]).reset_index(
                drop=True
            )
            late_orders_df.index += 1

            # Format date to BR standard
            if "Data de entrega" in late_orders_df.columns:
                late_orders_df["Data de entrega"] = pd.to_datetime(
                    late_orders_df["Data de entrega"]
                ).dt.strftime("%d/%m/%Y")

            late_orders_html = late_orders_df.to_html(col_space=50, justify="center")

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
                to_address = self.override_to or supplier_data["email"]
                logger.info(
                    f"Enviando e-mail de pedidos em atraso para: {supplier_data['name']} ({to_address})"
                )
                self._send_request(
                    email_to=to_address,
                    subject=f"Pedidos Atrasados {supplier_data['name']}",
                    body_html=email_body,
                    cc_recipients=self.emails_cc,
                )
                time.sleep(1)  # Small delay to avoid rate limits

            except Exception as e:
                logger.error(
                    f"Falha ao enviar e-mail CORRETIVO para {supplier_data['name']}: {str(e)}"
                )
                error_entry = pd.DataFrame(
                    {
                        "Name": [supplier_data["name"]],
                        "Email": [supplier_data["email"]],
                        "Error": [str(e)],
                    }
                )
                error_log = pd.concat([error_log, error_entry], ignore_index=True)

        if not error_log.empty:
            log_file_path = os.path.join(folder_path, "error_corrective_log.xlsx")
            error_log.to_excel(log_file_path, index=False)

    def send_preventive_email(self, suppliers: dict) -> None:
        """
        Send Preventive Email using Microsoft Graph API
        """
        error_log = pd.DataFrame(columns=["Name", "Email", "Error"])

        folder_path: str = "tmp"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        for _, supplier_data in suppliers.items():
            if not supplier_data["preventive_orders"]:
                continue

            preventive_orders_df = pd.DataFrame(
                supplier_data["preventive_orders"]
            ).reset_index(drop=True)
            preventive_orders_df.index += 1

            # Format date to BR standard
            if "Data de entrega" in preventive_orders_df.columns:
                preventive_orders_df["Data de entrega"] = pd.to_datetime(
                    preventive_orders_df["Data de entrega"]
                ).dt.strftime("%d/%m/%Y")

            preventive_orders_html = preventive_orders_df.to_html(
                col_space=50, justify="center"
            )

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
                to_address = self.override_to or supplier_data["email"]
                logger.info(
                    f"Enviando e-mail preventivo para: {supplier_data['name']} ({to_address})"
                )
                self._send_request(
                    email_to=to_address,
                    subject=f"Confirmação de Pedidos {supplier_data['name']}",
                    body_html=email_body,
                    cc_recipients=self.emails_cc,
                )
                time.sleep(1)

            except Exception as e:
                logger.error(
                    f"Falha ao enviar e-mail PREVENTIVO para {supplier_data['name']}: {str(e)}"
                )
                error_entry = pd.DataFrame(
                    {
                        "Name": [supplier_data["name"]],
                        "Email": [supplier_data["email"]],
                        "Error": [str(e)],
                    }
                )
                error_log = pd.concat([error_log, error_entry], ignore_index=True)

        if not error_log.empty:
            log_file_path = os.path.join(folder_path, "error_preventive_log.xlsx")
            error_log.to_excel(log_file_path, index=False)
