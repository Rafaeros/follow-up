from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFrame,
    QMessageBox
)
from PySide6.QtCore import Qt


class ConfigTab(QWidget):
    """
    Aba de Configurações de Credenciais.
    """

    def __init__(self, config_manager=None, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)

        title = QLabel("Configurações de Credenciais")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #0f172a;")
        subtitle = QLabel(
            "Gerencie suas informações de acesso e chaves de API com segurança."
        )
        subtitle.setStyleSheet("color: #64748b;")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        form_frame = QFrame()
        form_frame.setStyleSheet(
            """
            QFrame {
                background-color: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
            }
        """
        )
        form_layout = QVBoxLayout(form_frame)
        form_layout.setContentsMargins(20, 20, 20, 20)
        form_layout.setSpacing(15)

        card_title = QLabel("Credenciais do Sistema")
        card_title.setStyleSheet("font-size: 16px; font-weight: bold; border: none;")
        form_layout.addWidget(card_title)

        grid = QGridLayout()
        grid.setSpacing(15)

        # Site credentials (CargaMaquina)
        lbl_user = QLabel("Login / Usuário")
        lbl_user.setStyleSheet("font-weight: bold; border: none;")
        self.input_user = QLineEdit()
        self.input_user.setPlaceholderText("Ex: admin_followup")
        self.input_user.setText(self.config_manager.get("username", ""))

        lbl_pass = QLabel("Senha")
        lbl_pass.setStyleSheet("font-weight: bold; border: none;")
        self.input_pass = QLineEdit()
        self.input_pass.setPlaceholderText("••••••••")
        self.input_pass.setEchoMode(QLineEdit.Password)
        self.input_pass.setText(self.config_manager.get("password", ""))

        # Outlook SMTP credentials
        lbl_outlook_email = QLabel("Outlook Email")
        lbl_outlook_email.setStyleSheet("font-weight: bold; border: none;")
        self.input_outlook_email = QLineEdit()
        self.input_outlook_email.setPlaceholderText("Ex: seu_email@outlook.com")
        self.input_outlook_email.setText(self.config_manager.get("outlook_email", ""))

        lbl_outlook_pass = QLabel("Outlook Password")
        lbl_outlook_pass.setStyleSheet("font-weight: bold; border: none;")
        self.input_outlook_pass = QLineEdit()
        self.input_outlook_pass.setPlaceholderText("••••••••")
        self.input_outlook_pass.setEchoMode(QLineEdit.Password)
        self.input_outlook_pass.setText(self.config_manager.get("outlook_password", ""))

        grid.addWidget(lbl_user, 0, 0)
        grid.addWidget(self.input_user, 1, 0)
        grid.addWidget(lbl_pass, 0, 1)
        grid.addWidget(self.input_pass, 1, 1)

        grid.addWidget(lbl_outlook_email, 2, 0)
        grid.addWidget(self.input_outlook_email, 3, 0)
        grid.addWidget(lbl_outlook_pass, 2, 1)
        grid.addWidget(self.input_outlook_pass, 3, 1)

        form_layout.addLayout(grid)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_cancel = QPushButton("Cancelar")
        self.btn_save = QPushButton("Salvar Alterações")
        self.btn_save.clicked.connect(self.save_credentials)
        self.btn_save.setObjectName("primary")

        btn_layout.addWidget(self.btn_cancel)
        btn_layout.addWidget(self.btn_save)

        form_layout.addLayout(btn_layout)
        layout.addWidget(form_frame)

        info_frame = QFrame()
        info_frame.setStyleSheet(
            """
            QFrame {
                background-color: #f5f3ff; /* Fundo roxo bem claro */
                border: 1px solid #ddd6fe;
                border-radius: 8px;
            }
            QLabel { border: none; background: transparent; }
        """
        )
        info_layout = QVBoxLayout(info_frame)

        info_title = QLabel("Dica de Segurança")
        info_title.setStyleSheet("color: #7609e8; font-weight: bold;")
        info_text = QLabel(
            "Recomendamos a troca periódica de suas credenciais de acesso. Evite utilizar senhas óbvias."
        )
        info_text.setWordWrap(True)
        info_text.setStyleSheet("color: #475569;")
        info_layout.addWidget(info_title)
        info_layout.addWidget(info_text)
        layout.addWidget(info_frame)
        layout.addStretch()

    def save_credentials(self):
        username = self.input_user.text().strip()
        password = self.input_pass.text().strip()
        outlook_email = self.input_outlook_email.text().strip()
        outlook_password = self.input_outlook_pass.text().strip()

        if not username or not password:
            QMessageBox.warning(self, "Campos Incompletos", "Por favor, preencha todos os campos de login do site.")
            return

        self.config_manager.set_session_config({
            "username": username,
            "password": password,
        })

        # Outlook is optional, but if any field is filled, require both
        if outlook_email or outlook_password:
            if not outlook_email or not outlook_password:
                QMessageBox.warning(self, "Campos Incompletos", "Preencha ambos os campos do Outlook ou deixe-os em branco.")
                return
            self.config_manager.set_outlook_config({
                "outlook_email": outlook_email,
                "outlook_password": outlook_password,
            })

        QMessageBox.information(self, "Sucesso", "Credenciais salvas com sucesso!")