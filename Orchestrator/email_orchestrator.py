from ..auth import AuthContext
from ..EmailHandler import EmailContext
from ..common import ConnectionInfo, DataGetEmails, DataSendEmail, DataFiltersEmails, DataDownloadAttachments
import pandas as pd

class Orchestrator_email:
    def __init__(self, conn_info: ConnectionInfo):
        self.app_gateway = conn_info.email_provider
        self.auth_context = AuthContext(conn_info)
        self.auth_context.authenticate()
        self.namespace = self.auth_context.get_namespace()
        self.application = self.auth_context.get_application()
        self.email_service = EmailContext(self.namespace, self.application)
        
    def get_emails(self, datagetemails: DataGetEmails, return_all_pages: bool = False):
        result = self.email_service.get_emails(datagetemails)
        if return_all_pages:
            df_emails = result['data']
            page_next = result['page_next']
            has_more = result['has_more']
            while has_more:
                datagetemails.page_next = page_next
                result_next = self.email_service.get_emails(datagetemails)
                df_next = result_next['data']
                df_emails = pd.concat([df_emails, df_next], ignore_index=True)
                has_more = result_next['has_more']
                page_next = result_next['page_next']
            result['data'] = df_emails
        return result
    
    
    def send_email(self, datasentemail: DataSendEmail):
        return self.email_service.send_email(datasentemail)
        
