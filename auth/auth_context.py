from .auth_outlook import AuthOutlook
from ..common import ConnectionInfo

class AuthContext:
    
    def __init__(self, conn_info: ConnectionInfo):
        self.conn_info = conn_info
        if self.conn_info.email_provider.upper() == 'OUTLOOK':
            self._auth_strategy = AuthOutlook(conn_info)
        elif self.conn_info.email_provider.upper() == 'GMAIL':
            raise NotImplementedError("AuthGmail no está implementado aún.")
        else:
            raise ValueError(f"Proveedor de correo no soportado: {self.conn_info.email_provider}")
    
    def authenticate(self):
        return self._auth_strategy.authenticate()
    
    def get_namespace(self):
        return self._auth_strategy.get_namespace
    
    def get_application(self):
        return self._auth_strategy.get_application
    
    