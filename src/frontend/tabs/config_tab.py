from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFrame,
    QMessageBox,
)
from PySide6.QtCore import Qt
from src.utils.ms_graph_auth import MSGraphAuth


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

        # Microsoft Graph API credentials
        lbl_ms_graph_client_id = QLabel("MS Graph Client ID")
        lbl_ms_graph_client_id.setStyleSheet("font-weight: bold; border: none;")
        self.input_ms_graph_client_id = QLineEdit()
        self.input_ms_graph_client_id.setPlaceholderText("Ex: a1b2c3d4-...")
        self.input_ms_graph_client_id.setText(
            self.config_manager.get(
                "ms_graph_client_id", "15cd9ff3-25a9-4fb8-9433-6201eef53878"
            )
        )

        lbl_ms_graph_tenant_id = QLabel("MS Graph Tenant ID")
        lbl_ms_graph_tenant_id.setStyleSheet("font-weight: bold; border: none;")
        self.input_ms_graph_tenant_id = QLineEdit()
        self.input_ms_graph_tenant_id.setPlaceholderText("Ex: common ou seu-tenant-id")
        self.input_ms_graph_tenant_id.setText(
            self.config_manager.get(
                "ms_graph_tenant_id", "62c4daa3-2df5-40eb-9aa4-a0d5708ee0e7"
            )
        )

        grid.addWidget(lbl_user, 0, 0)
        grid.addWidget(self.input_user, 1, 0)
        grid.addWidget(lbl_pass, 0, 1)
        grid.addWidget(self.input_pass, 1, 1)

        grid.addWidget(lbl_ms_graph_client_id, 2, 0)
        grid.addWidget(self.input_ms_graph_client_id, 3, 0)
        grid.addWidget(lbl_ms_graph_tenant_id, 2, 1)
        grid.addWidget(self.input_ms_graph_tenant_id, 3, 1)

        form_layout.addLayout(grid)

        # MS Graph Auth Buttons
        ms_auth_layout = QHBoxLayout()
        self.btn_ms_login = QPushButton("Autenticar no Outlook (Navegador)")
        self.btn_ms_login.clicked.connect(self.authenticate_ms_graph)
        self.btn_ms_login.setStyleSheet(
            "background-color: #2b579a; color: white; font-weight: bold;"
        )

        self.btn_ms_logout = QPushButton("Limpar Login")
        self.btn_ms_logout.clicked.connect(self.logout_ms_graph)

        ms_auth_layout.addWidget(self.btn_ms_login)
        ms_auth_layout.addWidget(self.btn_ms_logout)
        form_layout.addLayout(ms_auth_layout)
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

        info_title = QLabel("Instruções de Login Outlook")
        info_title.setStyleSheet("color: #7609e8; font-weight: bold;")
        info_text = QLabel(
            "1. Verifique se o Client ID e o Tenant ID estão corretos.\n"
            "2. Clique em 'Salvar Alterações' primeiro para atualizar as configurações.\n"
            "3. Clique em 'Autenticar no Outlook' para abrir o navegador e fazer login.\n"
            "4. Após o login bem-sucedido, o sistema estará pronto para enviar e-mails."
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
        ms_graph_client_id = self.input_ms_graph_client_id.text().strip()
        ms_graph_tenant_id = self.input_ms_graph_tenant_id.text().strip()

        if not username or not password:
            QMessageBox.warning(
                self,
                "Campos Incompletos",
                "Por favor, preencha todos os campos de login do site.",
            )
            return

        # Salva as credenciais do site
        self.config_manager.set_session_config(
            {
                "username": username,
                "password": password,
            }
        )

        # Salva credenciais do MS Graph
        self.config_manager.set_ms_graph_config(
            {
                "ms_graph_client_id": ms_graph_client_id,
                "ms_graph_tenant_id": ms_graph_tenant_id,
            }
        )

        QMessageBox.information(self, "Sucesso", "Credenciais salvas com sucesso!")

    def authenticate_ms_graph(self):
        """Abre o navegador para autenticação no MS Graph."""
        try:
            # Pega as configurações atuais (salvas ou dos campos)
            client_id = self.input_ms_graph_client_id.text().strip()
            tenant_id = self.input_ms_graph_tenant_id.text().strip()

            if not client_id or not tenant_id:
                QMessageBox.warning(
                    self, "Erro", "Client ID e Tenant ID são obrigatórios."
                )
                return

            auth = MSGraphAuth(client_id=client_id, tenant_id=tenant_id)
            auth.login()
            QMessageBox.information(
                self, "Sucesso", "Autenticação realizada com sucesso!"
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Erro de Autenticação", f"Falha ao autenticar: {str(e)}"
            )

    def logout_ms_graph(self):
        """Limpa o cache de login do MS Graph."""
        try:
            auth = MSGraphAuth()
            auth.logout()
            QMessageBox.information(
                self, "Sucesso", "Login do Outlook removido com sucesso."
            )
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao limpar login: {str(e)}")
