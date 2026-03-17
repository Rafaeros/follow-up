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
        self.excluded_suppliers: set[str] = set()
        self.setup_ui()

    def setup_ui(self):
        self.layout = QVBoxLayout(self)

        self.filter_widget = FilterWidget()
        self.table_widget = ReportTable()
        self.table_widget.supplier_removed.connect(self.on_supplier_removed)

        # Layout inferior
        self.bottom_layout = QHBoxLayout()

        # Checkboxes
        self.chk_email_preventivo = QCheckBox("Enviar Email Preventivo")
        self.chk_email_preventivo.setChecked(True)
        self.chk_email_mp_atrasada = QCheckBox("Enviar Email de MP atrasada")
        self.chk_email_mp_atrasada.setChecked(True)

        # Botão
        self.btn_send_emails = QPushButton("Avançar para Envio de E-mails")
        self.btn_send_emails.setObjectName("primary")
        self.btn_send_emails.setEnabled(False)
        self.btn_send_emails.clicked.connect(self.processar_envio)

        # Montagem do layout
        self.bottom_layout.addWidget(self.chk_email_preventivo)
        self.bottom_layout.addWidget(self.chk_email_mp_atrasada)

        self.btn_remove_supplier = QPushButton("Remover Fornecedor Selecionado")
        self.btn_remove_supplier.clicked.connect(self.remove_selected_supplier)
        self.bottom_layout.addWidget(self.btn_remove_supplier)

        self.bottom_layout.addStretch()

        self.bottom_layout.addWidget(self.btn_send_emails)

        self.layout.addWidget(self.filter_widget)
        self.layout.addWidget(self.table_widget)
        self.layout.addLayout(self.bottom_layout)

        self.filter_widget.search_clicked.connect(self.start_report_scraping)

    def start_report_scraping(
        self, init_date: str, end_date: str, download_suppliers: bool
    ):
        # Resetar exclusões quando puxar um novo relatório (permite rodar várias vezes)
        self.excluded_suppliers.clear()

        # Verificar se fornecedores.json existe se não for atualizar
        if not download_suppliers:
            fornecedores_path = os.path.join("tmp", "fornecedores.json")
            if not os.path.exists(fornecedores_path):
                QMessageBox.warning(
                    self, "Aviso", "Arquivo fornecedores.json não encontrado. Ative a opção de atualizar fornecedores."
                )
                return

        self.filter_widget.btn_search.setEnabled(False)
        self.filter_widget.btn_search.setText("Buscando Dados...")
        self.btn_send_emails.setEnabled(False)

        # Delete existing report file to force regeneration
        dados_dir = os.path.join("tmp", "dados")
        if os.path.exists(dados_dir):
            for f in os.listdir(dados_dir):
                if f.endswith(".json") and "unified_report" in f:
                    os.remove(os.path.join(dados_dir, f))

        self.worker = ScraperWorker(
            self.session_manager, init_date, end_date, download_suppliers
        )
        self.worker.finished_success.connect(self.on_scraping_success)
        self.worker.finished_error.connect(self.on_scraping_error)
        self.worker.start()

    def on_scraping_success(self, unified_data: list):
        # Resetar exclusões quando um novo relatório é carregado
        self.excluded_suppliers.clear()

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

        # Se o usuário removeu fornecedores na tabela, manter o conjunto atualizado
        # para quando o envio de emails for acionado.

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

        # Process suppliers
        suppliers = {}
        for item in report_data:
            supplier_name = item["supplier_name"].strip()

            # Caso o usuário tenha removido esse fornecedor na tabela, ignorar
            if supplier_name in self.excluded_suppliers:
                continue

            email_str = item.get("email", "").strip()
            if not email_str or email_str == "-":
                continue  # Skip if no email
            # Use the first email if multiple
            email = email_str.split("; ")[0]

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
            override_to=None,
        )
        self.email_worker.finished_success.connect(self.on_email_success)
        self.email_worker.finished_error.connect(self.on_email_error)
        self.email_worker.progress_status.connect(self.update_email_status)
        self.email_worker.start()

    def update_email_status(self, status: str):
        self.btn_send_emails.setText(status)

    def on_supplier_removed(self, supplier_name: str):
        """Atualiza o estado interno quando um fornecedor for removido da tabela."""
        self.excluded_suppliers.add(supplier_name)

    def remove_selected_supplier(self):
        """Remove o fornecedor da linha selecionada da tabela e marca para não envio."""
        linha = self.table_widget.currentRow()
        if linha < 0:
            QMessageBox.warning(
                self,
                "Aviso",
                "Selecione uma linha na tabela antes de remover o fornecedor.",
            )
            return

        fornecedor = self.table_widget.item(linha, 2).text()
        self.table_widget.remover_fornecedor_selecionado()
        QMessageBox.information(
            self,
            "Removido",
            f"Fornecedor '{fornecedor}' removido. Ele não será incluído no envio de e-mails.",
        )

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
