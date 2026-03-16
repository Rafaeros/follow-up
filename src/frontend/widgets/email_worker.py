import logging
from PySide6.QtCore import QThread, Signal
from src.utils.email_sender import EmailSender

logger = logging.getLogger(__name__)


class EmailWorker(QThread):
    """
    Worker to handle asynchronous email sending for Microsoft Graph API.
    """

    finished_success = Signal()
    finished_error = Signal(str)
    progress_status = Signal(str)

    def __init__(
        self,
        suppliers: dict,
        send_preventive: bool,
        send_corrective: bool,
        override_to: str = None,
    ):
        super().__init__()
        self.suppliers = suppliers
        self.send_preventive = send_preventive
        self.send_corrective = send_corrective
        self.override_to = override_to

    def run(self):
        try:
            self.progress_status.emit("Conectando com Microsoft...")

            email_sender = EmailSender(override_to=self.override_to, disable_cc=False)

            # Explicitly trigger authentication to verify browser login
            try:
                self.progress_status.emit("Aguardando login no navegador...")
                email_sender.auth.get_access_token()
            except Exception as auth_error:
                error_msg = f"Falha na Autenticação OAuth2: {str(auth_error)}\n\nCertifique-se de que o navegador abriu e você concluiu o login."
                self.finished_error.emit(error_msg)
                return

            if self.send_corrective:
                self.progress_status.emit("Enviando e-mails em atraso...")
                email_sender.send_corrective_email(self.suppliers)

            if self.send_preventive:
                self.progress_status.emit("Enviando e-mails preventivos...")
                email_sender.send_preventive_email(self.suppliers)

            self.finished_success.emit()

        except Exception as e:
            logger.error(f"Erro no EmailWorker: {str(e)}")
            self.finished_error.emit(f"Erro inesperado no envio: {str(e)}")
