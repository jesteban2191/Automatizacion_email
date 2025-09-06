from ..common import ConnectionInfo
from .auth_interface import AuthInterface
import win32com.client
import pywintypes

class AuthOutlook(AuthInterface):
    
    def __init__(self, conn_info: ConnectionInfo):
        if conn_info.email_provider.upper() != 'OUTLOOK':
            raise ValueError("El email_provider debe ser 'outlook' para AuthOutlook")
        
        self._application = 'Outlook.Application'
        self._protocol = 'MAPI'
        
    
    def authenticate(self):
        try:
            self._outlook = win32com.client.Dispatch(self._application)
            self._namespace = self._outlook.GetNamespace(self._protocol)
        except ImportError as e:
            raise RuntimeError("No se encontró la librería win32com. Instálala con 'pip install pywin32'.") from e
        except pywintypes.com_error as e:
            raise RuntimeError("No se pudo iniciar Outlook o acceder al namespace. Verifica que Outlook esté instalado y configurado correctamente.") from e
        except AttributeError as e:
            raise RuntimeError("Error al acceder a los métodos de Outlook. Puede que la automatización COM esté fallando.") from e
        
        
    @property
    def get_namespace(self):
        return self._namespace
    
    
