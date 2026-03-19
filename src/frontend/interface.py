from PySide6.QtWidgets import QMainWindow, QTabWidget
from src.frontend.tabs.main_tab import MainTab
from src.frontend.tabs.config_tab import ConfigTab
from src.frontend.tabs.emails_tab import EmailTab


class Interface(QMainWindow):
    def __init__(self, config_manager, session_manager):
        super().__init__()
        self.config_manager = config_manager
        self.session_manager = session_manager
        self.setWindowTitle("Lanx Follow-Up")
        self.resize(1024, 768)
        self.setup_ui()

    def setup_ui(self):
        self.tabs = QTabWidget()
        self.main_tab = MainTab(parent=self)
        self.config_tab = ConfigTab(config_manager=self.config_manager, parent=self)
        self.emails_tab = EmailTab(parent=self)
        self.tabs.addTab(self.main_tab, "Relatórios")
        self.tabs.addTab(self.emails_tab, "Emails em Cópia")
        self.tabs.addTab(self.config_tab, "Configurações")
        self.setCentralWidget(self.tabs)
