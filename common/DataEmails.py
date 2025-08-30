from pydantic import BaseModel, field_validator
from typing import Optional
from enum import Enum
from datetime import datetime
import re
from ..common import ConnectionInfo

class OutlookStandarFolders(int, Enum):
    INBOX = 6  # Inbox
    OUTBOX = 3  # Outbox
    SENT_ITEMS = 4  # Sent Items
    DELETED_ITEMS = 5  # Deleted Items
    DRAFTS = 16  # Drafts
    JUNK_EMAIL = 14  # Junk Email
    CALENDAR = 9  # Calendar
    CONTACTS = 10  # Contacts
    TASKS = 12  # Tasks
    JOURNAL = 13  # Journal
    ROOT = 0  # Root
    
class IMPORTANCEEMAIL(int, Enum):
    LOW = 0
    NORMAL = 1
    HIGH = 2

class LOGICOPERATOR(str, Enum):
    AND = 'AND'
    OR = 'OR'

class SUBJECTPREFIX(str, Enum):
    RE = 'Re:'
    FWD = 'Fwd:'
    RV = 'RV:'
    TR = 'Tr:'
    RES = 'Res:'
    ENC = 'Enc:'

class QUERYDASL(str, Enum):
    SUBJECT = 'urn:schemas:mailheader:subject' # Asunto del correo
    BODY_TEXT = 'urn:schemas:httpmail:textdescription' #Cuerpo del correo en texto plano
    BODY_HTML = 'urn:schemas:httpmail:htmldescription' #Cuerpo del correo en HTML
    SENDER_EMAIL = 'urn:schemas:httpmail:fromemail' # Email del remitente
    RECIPIENT_EMAIL = 'urn:schemas:httpmail:to' # Email del destinatario
    CC_EMAIL = 'urn:schemas:httpmail:cc' # Email destinatarios en copia
    BCC_EMAIL = 'urn:schemas:httpmail:bcc' # Email destinatarios en copia oculta
    SENDER_NAME = 'urn:schemas:httpmail:fromname' # Nombre del remitente
    RECIPIENT_NAME = 'urn:schemas:httpmail:toname' # Nombre del destinatario
    CC_NAME = 'urn:schemas:httpmail:ccname' # Nombre destinatarios en copia
    BCC_NAME = 'urn:schemas:httpmail:bccname' # Nombre destinatarios en copia oculta
    IMPORTANCE = 'urn:schemas:httpmail:importance' # Importancia del correo (1=Alta, 2=Normal, 3=Baja)
    HAS_ATTACHMENTS = 'urn:schemas:httpmail:hasattachment' # Si el correo tiene adjuntos
    IS_READ = 'urn:schemas:httpmail:read' # Si el correo ha sido leído
    RECEIVED_TIME = 'urn:schemas:httpmail:datereceived' # Fecha y hora de recepción
    SENT_ON = 'urn:schemas:httpmail:date' # Fecha y hora de envío
    DELIVERED_TIME = 'urn:schemas:httpmail:deliverytime' # Fecha y hora de entrega
    CONVERSATION_TOPIC = 'urn:schemas:httpmail:thread-topic' # Asunto del hilo de conversación
    REFERENCEID = 'urn:schemas:httpmail:references' # Message-ID de los emails referenciados (header References)
    MESSAGE_ID = 'urn:schemas:httpmail:messageid' # Message-ID del email (header)
    SUBJECT_PREFIX = 'urn:schemas:httpmail:subjectprefix' # Prefijo del asunto (Re:, Fwd:, etc.)
    

class DataFiltersEmails(BaseModel):
    subject: Optional[str] = None
    body: Optional[str] = None
    sender: Optional[list[str]] = None
    recipient: Optional[list[str]] = None
    sender_email: Optional[list[str]] = None  # Lista de emails del remitente
    recipient_email: Optional[list[str]] = None  # Lista de emails del destinatario
    cc_email: Optional[list[str]] = None  # Lista de emails en copia
    bcc_email: Optional[list[str]] = None  # Lista de emails en copia oculta
    cc: Optional[list[str]] = None  # Lista de nombres en copia
    bcc: Optional[list[str]] = None  # Lista de nombres en copia oculta
    has_attachments: Optional[bool] = None
    is_read: Optional[bool] = None
    received_after: Optional[str] = None  # ISO format date string
    received_before: Optional[str] = None  # ISO format date string
    conversation_topic: Optional[str] = None  # Corresponde al asunto del hilo de conversación (conversation topic)
    referenceid: Optional[list[str]] = None  # Corresponde al Message-ID de los emails referenciados (header References)
    msg_id: Optional[str] = None  # Corresponde al Message-ID del email (header)
    importance_email: Optional[IMPORTANCEEMAIL] = None  # IMPORTANCEEMAIL Enum
    logic_operator: Optional[LOGICOPERATOR] = LOGICOPERATOR.AND  # 'AND' o 'OR'
    logic_operator_between_senders: Optional[LOGICOPERATOR] = LOGICOPERATOR.OR  # 'AND' o 'OR' entre los remitentes
    logic_operator_between_recipients: Optional[LOGICOPERATOR] = LOGICOPERATOR.OR  # 'AND' o 'OR' entre los destinatarios
    subject_prefix: Optional[SUBJECTPREFIX] = None  # Prefijo del asunto (Re:, Fwd:, etc.)

    @field_validator('sender', 'recipient', mode='before')
    def validate_email_format(cls, v):
        if v is None:
            return v
        email_regex = r'^[\w\.-]+@[\w\.-]+\.[a-zA-Z\.]+$'
        if isinstance(v, list):
            for email in v:
                if not re.match(email_regex, email):
                    raise ValueError(f"Email inválido: {email}")
        else:
            if not re.match(email_regex, v):
                raise ValueError(f"Email inválido: {v}")
        return v
        
        
    @field_validator('received_after', 'received_before', mode='before')
    def validate_date_format(cls, v):
        if v is None:
            return v
        try: 
            dt = datetime.strptime(v, '%d/%m/%Y %H:%M:%S')
        except Exception:
            try:
                dt = datetime.strptime(v, '%d/%m/%Y')
            except Exception:
                raise ValueError(f"El formato de fecha debe ser 'dd/mm/yyyy hh:mm:ss'. Valor recibido: {v}")
        return dt  # ¡Siempre retorna el valor!
    

class DataGetEmails(BaseModel):
    standard_folder: Optional[OutlookStandarFolders] = OutlookStandarFolders.INBOX
    custom_folder: Optional[str] = None  # Esta es la ruta completa de la carpeta personalizada
    max_emails: Optional[int] = 100  # Número máximo de emails a obtener
    filters: Optional[DataFiltersEmails] = None
    mark_as_read: Optional[bool] = False  # Marcar los emails obtenidos como leídos
    page_next: Optional[int] = None # Página siguiente a obtener, si es None, se obtiene la primera página
    
    @field_validator('custom_folder', mode='before')
    def validate_custom_folder(cls, v):
        if v is None:
            return v
        # Aquí puedes agregar validaciones específicas para la carpeta personalizada
        v = v.replace('\\', '/')
        folder_regex = r'^[\w/]+$'
        if not re.match(folder_regex, v):
            raise ValueError("La ruta solo puede contener letras, números, guion bajo y '/'. Para separar subcarpetas se debe usar '/'.")
        return v
    


