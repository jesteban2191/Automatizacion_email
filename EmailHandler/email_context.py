from .outlook_email import OutlookEmail
from ..common import DataGetEmails, DataFiltersEmails
import win32com.client

class EmailContext:
    def __init__(self, conn):
        if hasattr(conn, "_oleobj_") and hasattr(conn, "GetDefaultFolder"):
            self.email_handler = OutlookEmail(conn)
        else:
            raise ValueError("El objeto de conexión no es válido. Debe ser un objeto Outlook autenticado.")
        
        
    def get_emails(self, datagetemails: DataGetEmails):
        return self.email_handler.get_emails(datagetemails)
    
    def create_query(self, datafiltersemails: DataFiltersEmails):
        return self.email_handler.create_query(datafiltersemails)
    
    