# src/gui/components/report_table.py
from PySide6.QtWidgets import QTableWidget, QHeaderView, QTableWidgetItem, QMenu
from PySide6.QtGui import QColor, QAction
from PySide6.QtCore import Qt, Signal
from datetime import datetime

class ReportTable(QTableWidget):
    """
    Responsabilidade: Configurar e gerenciar a exibição da tabela de relatórios.
    """

    supplier_removed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        self.setColumnCount(5)
        # Adicionado o "Cód." e mudado para "Previsão de Entrega"
        self.setHorizontalHeaderLabels(["Neg.", "Cód.", "Fornecedor", "Previsão de Entrega", "Status"])
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectRows)
        self.setEditTriggers(QTableWidget.NoEditTriggers)
        
        # --- Configuração do Menu de Contexto (Botão Direito) ---
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.abrir_menu_contexto)

    def abrir_menu_contexto(self, position):
        """Abre o menu para deletar quando clica com o botão direito."""
        menu = QMenu()
        acao_remover_linha = QAction("Remover Apenas Este Pedido", self)
        acao_remover_fornecedor = QAction("Remover TODOS os Pedidos deste Fornecedor", self)
        
        acao_remover_linha.triggered.connect(self.remover_linha_selecionada)
        acao_remover_fornecedor.triggered.connect(self.remover_fornecedor_selecionado)
        
        menu.addAction(acao_remover_linha)
        menu.addAction(acao_remover_fornecedor)
        
        # Exibe o menu na posição exata do mouse
        menu.exec(self.viewport().mapToGlobal(position))

    def remover_linha_selecionada(self):
        """Remove apenas a linha que estava selecionada."""
        linha_atual = self.currentRow()
        if linha_atual >= 0:
            self.removeRow(linha_atual)

    def keyPressEvent(self, event):
        """Permite deletar o fornecedor pressionando a tecla Delete."""
        from PySide6.QtCore import Qt

        if event.key() == Qt.Key_Delete:
            self.remover_fornecedor_selecionado()
        else:
            super().keyPressEvent(event)

    def remover_fornecedor_selecionado(self):
        """Remove todas as linhas da tabela que tiverem o mesmo nome de fornecedor."""
        linha_atual = self.currentRow()
        if linha_atual >= 0:
            # O nome do fornecedor está na coluna de índice 2
            nome_fornecedor = self.item(linha_atual, 2).text()

            # Emitir sinal para que a lógica externa (MainTab) remova este fornecedor da lista de envios
            self.supplier_removed.emit(nome_fornecedor)

            # Loop reverso: deletamos de baixo para cima para não quebrar os índices
            for row in range(self.rowCount() - 1, -1, -1):
                if self.item(row, 2).text() == nome_fornecedor:
                    self.removeRow(row)

    def formatar_data_br(self, data_str: str) -> str:
        """Converte YYYY-MM-DD para DD/MM/YYYY."""
        if not data_str or data_str == "None":
            return "-"
        try:
            data_limpa = data_str.split(" ")[0] 
            dt_obj = datetime.strptime(data_limpa, "%Y-%m-%d")
            return dt_obj.strftime("%d/%m/%Y")
        except Exception:
            return data_str

    def populate_data(self, unified_data: list):
        """Popula a tabela e aplica as cores de atraso/prazo."""
        self.setRowCount(0) 
        row_idx = 0

        for supplier in unified_data:
            fornecedor_nome = supplier.get("supplier_name", "-")

            # Popula Pedidos Atrasados
            for order in supplier.get("late_orders", []):
                self.insertRow(row_idx)
                self.setItem(row_idx, 0, QTableWidgetItem(str(order.get("Neg.", ""))))
                self.setItem(row_idx, 1, QTableWidgetItem(str(order.get("Cod.", ""))))
                self.setItem(row_idx, 2, QTableWidgetItem(fornecedor_nome))
                self.setItem(row_idx, 3, QTableWidgetItem(self.formatar_data_br(order.get("Data de entrega"))))
                
                status_item = QTableWidgetItem(f"Atrasado ({order.get('Dias de Atraso')} dias)")
                status_item.setForeground(QColor("red"))
                self.setItem(row_idx, 4, status_item)
                row_idx += 1

            # Popula Pedidos no Prazo
            for order in supplier.get("future_orders", []):
                self.insertRow(row_idx)
                self.setItem(row_idx, 0, QTableWidgetItem(str(order.get("Neg.", ""))))
                self.setItem(row_idx, 1, QTableWidgetItem(str(order.get("Cod.", ""))))
                self.setItem(row_idx, 2, QTableWidgetItem(fornecedor_nome))
                self.setItem(row_idx, 3, QTableWidgetItem(self.formatar_data_br(order.get("Data de entrega"))))
                
                status_item = QTableWidgetItem("No Prazo")
                status_item.setForeground(QColor("green"))
                self.setItem(row_idx, 4, status_item)
                row_idx += 1