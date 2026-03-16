from PySide6.QtWidgets import QWidget, QHBoxLayout, QPushButton, QCheckBox

class ActionFooter(QWidget):
    """
    Responsabilidade: Conter os botões de ação globais da aba.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Checkboxes
        self.chk_preventivo = QCheckBox("Enviar Email Preventivo")
        self.chk_mp_atrasada = QCheckBox("Enviar Email de MP atrasada")

        # botão
        self.btn_email = QPushButton("Avançar para Envio de E-mails")
        self.btn_email.setObjectName("primary")
        self.btn_email.setMinimumHeight(45)
        self.btn_email.setMinimumWidth(220)

        layout.addWidget(self.chk_preventivo)
        layout.addWidget(self.chk_mp_atrasada)

        layout.addStretch()

        layout.addWidget(self.btn_email)