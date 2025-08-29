from pydantic import BaseModel, model_validator
from typing import Optional
from enum import Enum

class PermisosGmail(str, Enum):
    EMAIL_SOLO_LECTURA = "https://www.googleapis.com/auth/gmail.readonly"
    EMAIL_ESCRITURA_LECTURA = "https://www.googleapis.com/auth/gmail.modify"
    EMAIL_SOLO_ENVIAR = "https://www.googleapis.com/auth/gmail.send"
    EMAIL_CREAR_ENVIAR = "https://www.googleapis.com/auth/gmail.compose"
    EMAIL_INSERTAR = "https://www.googleapis.com/auth/gmail.insert"
    EMAIL_GESTIONAR_LABELS = "https://www.googleapis.com/auth/gmail.labels"
    EMAIL_CONFIG_BASICA = "https://www.googleapis.com/auth/gmail.settings.basic"
    EMAIL_CONFIG_DELEGACION = "https://www.googleapis.com/auth/gmail.settings.sharing"
    EMAIL_ADMIN = "https://mail.google.com/"
    CALENDAR_SOLO_LECTURA = "https://www.googleapis.com/auth/calendar.readonly"
    CALENDAR_LECTURA_ESCRITURA = "https://www.googleapis.com/auth/calendar"
    CONTACTO_SOLO_LECTURA = "https://www.googleapis.com/auth/contacts.readonly"
    CONTACTO_LECTURA_ESCRITURA = "https://www.googleapis.com/auth/contacts"
    DRIVE_SOLO_LECTURA = "https://www.googleapis.com/auth/drive.readonly"
    DRIVE_LECTURA_ESCRITURA = "https://www.googleapis.com/auth/drive"
    DRIVE_FILE_APP = "https://www.googleapis.com/auth/drive.file"
    
class CredentialsInfoPath(BaseModel):
    token_path: Optional[str] = None
    credentials_path: Optional[str] = None
    token: Optional[str] = None

class ConnectionInfo(BaseModel):
    email_provider: str  # 'gmail' o 'outlook'
    scopes: Optional[list[PermisosGmail]] = None
    cred_info: Optional[CredentialsInfoPath] = None
    
    @model_validator(mode='after')
    def check_info_email(self):
        provider = self.email_provider.upper()
        cred_info = self.cred_info
        error_msg = []
        if provider not in ['GMAIL', 'OUTLOOK']:
            raise ValueError("- El email_provider debe ser 'gmail' o 'outlook'")
        if provider == 'GMAIL':
            if not cred_info or not(cred_info.credentials_path and cred_info.token_path and cred_info.token):
                error_msg.append("- Para 'gmail', se debe enviar la ruta del archivo de credenciales, ruta del archivo del token o el token en sí. Asegurarse de enviar por lo menos uno de estos parámetros.")
            if not cred_info.scopes:
                error_msg.append("- Para 'gmail', se deben enviar los permisos que se quieren aplicar a la aplicación.")

        elif provider == 'OUTLOOK':
            pass

        if error_msg:
            raise ValueError("Error en la Información de conexión:\n" + "\n".join(error_msg))

        return self