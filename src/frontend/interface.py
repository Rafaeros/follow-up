import asyncio
from PySide6.QtWidgets import QMainWindow, QTabWidget
from PySide6.QtGui import QCloseEvent
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

    def closeEvent(self, event: QCloseEvent):
        """Fecha a sessão HTTP de forma segura ao encerrar a UI."""
        session = getattr(self.session_manager, "session", None)
        if session and not session.closed:
            try:
                loop = asyncio.get_event_loop()
                if loop.is_running():
                    # Agendamos o fechamento para o loop existente (qasync)
                    loop.create_task(session.close())
                else:
                    loop.run_until_complete(session.close())
            except RuntimeError:
                # Caso não exista loop, use um temporário
                loop = asyncio.new_event_loop()
                loop.run_until_complete(session.close())
                loop.close()

        event.accept()