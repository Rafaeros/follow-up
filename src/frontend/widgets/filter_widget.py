# src/gui/components/filter_widget.py
from PySide6.QtWidgets import QWidget, QHBoxLayout, QLabel, QDateEdit, QPushButton
from PySide6.QtCore import Signal, QDate

class FilterWidget(QWidget):
    """
    Responsabilidade: Exibir inputs de data e emitir um sinal quando o usuário quiser buscar.
    """

    search_clicked = Signal(str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Data Inicial com Calendário
        self.start_date = QDateEdit(QDate.currentDate().addDays(-15))
        self.start_date.setCalendarPopup(True)  # <--- Habilita o calendário visual
        self.start_date.setDisplayFormat("dd/MM/yyyy") # <--- Força o padrão BR na tela

        # Data Final com Calendário
        self.end_date = QDateEdit(QDate.currentDate().addDays(30))
        self.end_date.setCalendarPopup(True)  # <--- Habilita o calendário visual
        self.end_date.setDisplayFormat("dd/MM/yyyy") # <--- Força o padrão BR na tela

        self.btn_search = QPushButton("Puxar Relatório")
        self.btn_search.setObjectName("primary")

        # Quando clicar, dispara a função local
        self.btn_search.clicked.connect(self.emit_search)

        layout.addWidget(QLabel("Data Inicial:"))
        layout.addWidget(self.start_date)
        layout.addWidget(QLabel("Data Final:"))
        layout.addWidget(self.end_date)
        layout.addWidget(self.btn_search)
        layout.addStretch()

    def emit_search(self):
        start = self.start_date.date().toString("dd/MM/yyyy")
        end = self.end_date.date().toString("dd/MM/yyyy")
        self.search_clicked.emit(start, end)