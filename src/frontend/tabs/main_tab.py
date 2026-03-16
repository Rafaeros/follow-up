# src/frontend/tabs/main_tab.py
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QMessageBox,
    QPushButton,
    QCheckBox,
)

from src.frontend.widgets.filter_widget import FilterWidget
from src.frontend.widgets.report_table import ReportTable
from src.frontend.widgets.scraper_worker import ScraperWorker
from src.utils.email_sender import EmailSender
from src.core.config import ConfigManager
import json
import os


class MainTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.session_manager = parent.session_manager
        self.setup_ui()

    def setup_ui(self):
        self.layout = QVBoxLayout(self)

        self.filter_widget = FilterWidget()
        self.table_widget = ReportTable()

        # Layout inferior
        self.bottom_layout = QHBoxLayout()

        # Checkboxes
        self.chk_email_preventivo = QCheckBox("Enviar Email Preventivo")
        self.chk_email_mp_atrasada = QCheckBox("Enviar Email de MP atrasada")
        self.chk_debug = QCheckBox("Debug (enviar para email do config)")

        # Botão
        self.btn_send_emails = QPushButton("Avançar para Envio de E-mails")
        self.btn_send_emails.setObjectName("primary")
        self.btn_send_emails.setEnabled(False)
        self.btn_send_emails.clicked.connect(self.processar_envio)

        # Montagem do layout
        self.bottom_layout.addWidget(self.chk_email_preventivo)
        self.bottom_layout.addWidget(self.chk_email_mp_atrasada)
        self.bottom_layout.addWidget(self.chk_debug)

        self.bottom_layout.addStretch()

        self.bottom_layout.addWidget(self.btn_send_emails)

        self.layout.addWidget(self.filter_widget)
        self.layout.addWidget(self.table_widget)
        self.layout.addLayout(self.bottom_layout)

        self.filter_widget.search_clicked.connect(self.start_report_scraping)

    def start_report_scraping(
        self, init_date: str, end_date: str, download_suppliers: bool
    ):
        self.filter_widget.btn_search.setEnabled(False)
        self.filter_widget.btn_search.setText("Buscando Dados...")
        self.btn_send_emails.setEnabled(False)

        self.worker = ScraperWorker(
            self.session_manager, init_date, end_date, download_suppliers
        )
        self.worker.finished_success.connect(self.on_scraping_success)
        self.worker.finished_error.connect(self.on_scraping_error)
        self.worker.start()

    def on_scraping_success(self, unified_data: list):
        self.table_widget.populate_data(unified_data)
        self.reset_ui_state()

        if unified_data:
            self.btn_send_emails.setEnabled(True)
        else:
            QMessageBox.information(
                self, "Aviso", "Nenhum dado encontrado no período selecionado."
            )

    def on_scraping_error(self, error_msg: str):
        QMessageBox.critical(
            self, "Erro de Conexão", f"Falha ao gerar o relatório:\n{error_msg}"
        )
        self.reset_ui_state()

    def reset_ui_state(self):
        self.filter_widget.btn_search.setEnabled(True)
        self.filter_widget.btn_search.setText("Puxar Relatório")

    def processar_envio(self):
        """Coleta os dados do unified_report.json para o envio"""
        # IMPORT WORKER HERE TO AVOID CIRCULAR IMPORTS
        from src.frontend.widgets.email_worker import EmailWorker

        # Find the latest unified_report.json
        dados_dir = os.path.join("tmp", "dados")
        if not os.path.exists(dados_dir):
            QMessageBox.warning(self, "Erro", "Diretório tmp/dados não encontrado.")
            return

        # Support file names like 'unified_report.json' and '<date>_unified_report.json'
        files = [
            f
            for f in os.listdir(dados_dir)
            if f.endswith(".json") and "unified_report" in f
        ]
        if not files:
            QMessageBox.warning(
                self, "Erro", "Arquivo unified_report.json não encontrado."
            )
            return

        # Sort by modification time, get the latest
        files.sort(
            key=lambda x: os.path.getmtime(os.path.join(dados_dir, x)), reverse=True
        )
        report_path = os.path.join(dados_dir, files[0])

        with open(report_path, "r", encoding="utf-8") as f:
            report_data = json.load(f)

        enviar_preventivo = self.chk_email_preventivo.isChecked()
        enviar_mp_atrasada = self.chk_email_mp_atrasada.isChecked()

        if not enviar_preventivo and not enviar_mp_atrasada:
            QMessageBox.warning(
                self, "Aviso", "Selecione pelo menos um tipo de email para enviar."
            )
            return

        # Default recipient for testing (sent to Guilherme and CC)
        forced_debug_email = "guilherme.silva@lanxcables.com.br"

        # Process suppliers
        suppliers = {}
        for item in report_data:
            supplier_name = item["supplier_name"].strip()
            # Always use the forced debug email instead of item['email']
            email = forced_debug_email

            if not email or email == "-":
                continue  # Skip if no email

            suppliers[supplier_name] = {
                "name": supplier_name,
                "email": email,
                "late_orders": item["late_orders"],
                "preventive_orders": item[
                    "future_orders"
                ],  # Map future_orders to preventive_orders
            }

        if not suppliers:
            QMessageBox.information(
                self,
                "Aviso",
                "Nenhum fornecedor elegível para envio de e-mail encontrado.",
            )
            return

        # Prepare for send
        self.btn_send_emails.setEnabled(False)
        self.btn_send_emails.setText("Enviando E-mails...")
        self.filter_widget.btn_search.setEnabled(False)

        # Send emails using worker to avoid blocking UI
        self.email_worker = EmailWorker(
            suppliers=suppliers,
            send_preventive=enviar_preventivo,
            send_corrective=enviar_mp_atrasada,
            override_to=forced_debug_email,
        )
        self.email_worker.finished_success.connect(self.on_email_success)
        self.email_worker.finished_error.connect(self.on_email_error)
        self.email_worker.progress_status.connect(self.update_email_status)
        self.email_worker.start()

    def update_email_status(self, status: str):
        self.btn_send_emails.setText(status)

    def on_email_success(self):
        self.btn_send_emails.setText("Avançar para Envio de E-mails")
        self.btn_send_emails.setEnabled(False)
        self.filter_widget.btn_search.setEnabled(True)

        # Clear the table as requested
        self.table_widget.populate_data([])

        QMessageBox.information(self, "Sucesso", "Emails processados com sucesso!")

    def on_email_error(self, error_msg: str):
        self.btn_send_emails.setText("Avançar para Envio de E-mails")
        self.btn_send_emails.setEnabled(True)
        self.filter_widget.btn_search.setEnabled(True)
        QMessageBox.critical(
            self, "Erro no Envio", f"Ocorreu um erro ao enviar e-mails:\n{error_msg}"
        )
