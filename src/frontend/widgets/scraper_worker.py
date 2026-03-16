# src/gui/workers/scraper_worker.py
import asyncio
from PySide6.QtCore import QThread, Signal
from src.core.session_manager import SessionManager
from src.core.scraper import ReportScraper # Ajuste o import conforme sua estrutura

class ScraperWorker(QThread):
    """
    Worker para rodar a extração de dados sem travar a GUI do PySide6.
    """
    finished_success = Signal(list)
    finished_error = Signal(str)

    def __init__(self, session_manager: SessionManager, init_date: str, end_date: str):
        super().__init__()
        self.session_manager = session_manager
        self.init_date = init_date
        self.end_date = end_date

    def run(self):
        try:
            # Função async interna para lidar com a event loop da thread
            async def fetch_data():
                # Garante que está logado antes de puxar
                if not self.session_manager.session:
                    await self.session_manager.login() 

                scraper = ReportScraper(self.session_manager)
                return await scraper.get_all_reports(self.init_date, self.end_date)

            # Cria um novo event loop para essa Thread específica
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
            # Executa a tarefa e pega o resultado
            result = loop.run_until_complete(fetch_data())
            loop.close()

            # Emite o sinal de sucesso com os dados retornados
            self.finished_success.emit(result)

        except Exception as e:
            self.finished_error.emit(str(e))