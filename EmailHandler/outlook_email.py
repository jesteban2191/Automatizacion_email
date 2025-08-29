from .email_interface import EmailInterface
from ..common import ConnectionInfo, DataGetEmails, DataFiltersEmails
import win32com.client

class OutlookEmail(EmailInterface):

    def __init__(self, conn: win32com.client.CDispatch):
        self._conn = conn # Este es el objeto Namespace autenticado
        
    
    def get_emails(self, datagetemails: DataGetEmails):
        
        
        #########################################################################################
        #####          Obtengo todos los parámetros para descagar los emails            #########
        #########################################################################################
        self.standard_folders = datagetemails.standard_folder
        self.custom_folder = datagetemails.custom_folder
        if self.custom_folder:
            self.custom_folder = self.custom_folder.split('/')
        self.max_emails = datagetemails.max_emails
        self.mark_as_read = datagetemails.mark_as_read
        self.subject = datagetemails.filters.subject
        self.sender = datagetemails.filters.sender
        self.recipient = datagetemails.filters.recipient
        self.body = datagetemails.filters.body
        self.has_attachments = datagetemails.filters.has_attachments
        self.is_read = datagetemails.filters.is_read
        self.received_after = datagetemails.filters.received_after
        self.received_before = datagetemails.filters.received_before
        self.conversation_topic = datagetemails.filters.conversation_topic
        self.msg_id = datagetemails.filters.msg_id
        
        #########################################################################################
        #####          Empiezo a crear el query para obtener los emails                 #########
        #########################################################################################

