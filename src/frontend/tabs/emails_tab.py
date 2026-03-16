
import json
import os
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QListWidget, QListWidgetItem, QMessageBox
from PySide6.QtCore import Qt


class EmailTab(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.emails_file = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'emails_cc.json')
        self.emails = self.load_emails()
        self.setup_ui()
        self.apply_theme()

    def setup_ui(self):
        layout = QVBoxLayout()

        # Input layout
        input_layout = QHBoxLayout()
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("Digite o email para cópia")
        self.add_button = QPushButton("Adicionar Email")
        self.add_button.setObjectName("primary")
        self.add_button.clicked.connect(self.add_email)
        input_layout.addWidget(self.email_input)
        input_layout.addWidget(self.add_button)

        # List widget
        self.email_list = QListWidget()
        self.email_list.setSelectionMode(QListWidget.MultiSelection)
        self.populate_list()

        # Delete button
        self.delete_button = QPushButton("Remover Selecionados")
        self.delete_button.setObjectName("primary")
        self.delete_button.clicked.connect(self.remove_selected_emails)

        layout.addLayout(input_layout)
        layout.addWidget(self.email_list)
        layout.addWidget(self.delete_button)

        self.setLayout(layout)

    def apply_theme(self):
        theme_path = os.path.join(os.path.dirname(__file__), '..', 'theme.qss')
        if os.path.exists(theme_path):
            with open(theme_path, 'r') as f:
                stylesheet = f.read()
            self.setStyleSheet(stylesheet)

    def load_emails(self):
        if os.path.exists(self.emails_file):
            try:
                with open(self.emails_file, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                return []
        return []

    def save_emails(self):
        with open(self.emails_file, 'w') as f:
            json.dump(self.emails, f, indent=4)

    def add_email(self):
        email = self.email_input.text().strip()
        if email and email not in self.emails:
            self.emails.append(email)
            self.save_emails()
            self.populate_list()
            self.email_input.clear()
        elif email in self.emails:
            QMessageBox.warning(self, "Aviso", "Email já adicionado.")
        else:
            QMessageBox.warning(self, "Aviso", "Digite um email válido.")

    def populate_list(self):
        self.email_list.clear()
        for email in self.emails:
            item = QListWidgetItem(email)
            self.email_list.addItem(item)

    def remove_selected_emails(self):
        selected_items = self.email_list.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Info", "Selecione emails para remover.")
            return
        for item in selected_items:
            email = item.text()
            if email in self.emails:
                self.emails.remove(email)
        self.save_emails()
        self.populate_list()
