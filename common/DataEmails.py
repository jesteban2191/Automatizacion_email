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
    
class OutlookStandardFoldersstr(str, Enum):
    INBOX = 'Bandeja de entrada'  # Inbox
    OUTBOX = 'Bandeja de salida'  # Outbox
    SENT_ITEMS = 'Elementos enviados'  # Sent Items
    DELETED_ITEMS = 'Elementos eliminados'  # Deleted Items
    DRAFTS = 'Borradores'  # Drafts
    JUNK_EMAIL = 'Correo no deseado'  # Junk Email
    CALENDAR = 'Calendario'  # Calendar
    CONTACTS = 'Contactos'  # Contacts
    TASKS = 'Tareas'  # Tasks
    JOURNAL = 'Diario'  # Journal
    ROOT = 'Raíz'  # Root
 
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
    
    #######################################################################
    #####             Campos DASL para mensajes IPM.Note              #####
    #######################################################################
    SUBJECT_IPM_NOTE = 'urn:schemas:httpmail:subject' # Asunto del correo
    BODY_TEXT_IPM_NOTE = 'urn:schemas:httpmail:textdescription' #Cuerpo del correo en texto plano
    BODY_HTML_IPM_NOTE = 'urn:schemas:httpmail:htmlbody' #Cuerpo del correo en HTML
    SENDER_EMAIL_IPM_NOTE = 'urn:schemas:httpmail:fromemail' # Email del remitente
    RECIPIENT_EMAIL_IPM_NOTE = 'urn:schemas:httpmail:to' # Email del destinatario
    CC_EMAIL_IPM_NOTE = 'urn:schemas:httpmail:cc' # Email destinatarios en copia
    BCC_EMAIL_IPM_NOTE = 'urn:schemas:httpmail:bcc' # Email destinatarios en copia oculta
    SENDER_NAME_IPM_NOTE = 'urn:schemas:httpmail:fromname' # Nombre del remitente
    RECIPIENT_NAME_IPM_NOTE = 'urn:schemas:httpmail:toname' # Nombre del destinatario
    CC_NAME_IPM_NOTE = 'urn:schemas:httpmail:ccname' # Nombre destinatarios en copia
    BCC_NAME_IPM_NOTE = 'urn:schemas:httpmail:bccname' # Nombre destinatarios en copia oculta
    IMPORTANCE_IPM_NOTE = 'urn:schemas:httpmail:importance' # Importancia del correo (1=Alta, 2=Normal, 3=Baja)
    HAS_ATTACHMENTS_IPM_NOTE = 'urn:schemas:httpmail:hasattachment' # Si el correo tiene adjuntos
    IS_READ_IPM_NOTE = 'urn:schemas:httpmail:read' # Si el correo ha sido leído
    RECEIVED_TIME_IPM_NOTE = 'urn:schemas:httpmail:datereceived' # Fecha y hora de recepción
    SENT_ON_IPM_NOTE = 'urn:schemas:httpmail:date' # Fecha y hora de envío
    DELIVERED_TIME_IPM_NOTE = 'urn:schemas:httpmail:deliverytime' # Fecha y hora de entrega
    CONVERSATION_TOPIC_IPM_NOTE = 'urn:schemas:httpmail:thread-topic' # Asunto del hilo de conversación
    REFERENCEID_IPM_NOTE = 'urn:schemas:httpmail:references' # Message-ID de los emails referenciados (header References)
    ID_IPM_NOTE = 'urn:schemas:httpmail:messageid' # Message-ID del email (header)
    SUBJECT_PREFIX_IPM_NOTE = 'urn:schemas:httpmail:subjectprefix' # Prefijo del asunto (Re:, Fwd:, etc.)
    MSGCLASS_IPM_NOTE = 'urn:schemas:httpmail:messageclass' # Clase del mensaje (IPM.Note, IPM.Appointment, etc.)
    
    ############################################################################
    ##### Campos DASL para mensajes IPM.Appointment o IPM.Schedule_Meeting #####
    ############################################################################
    SUBJECT_IPM_MEETING = 'urn:schemas:calendar:subject' # Asunto de la cita o reunión
    BODY_TEXT_IPM_MEETING = 'urn:schemas:calendar:body' # Cuerpo de la cita o reunión en texto plano
    LOCATION_IPM_MEETING = 'urn:schemas:calendar:location' # Ubicación de la cita o reunión
    START_TIME_IPM_MEETING = 'urn:schemas:calendar:starttime' # Fecha
    END_TIME_IPM_MEETING = 'urn:schemas:calendar:endtime' # Fecha de fin
    ORGANIZER_IPM_MEETING = 'urn:schemas:calendar:organizer' # Organizador de la cita o reunión
    DURATION_IPM_MEETING = 'urn:schemas:calendar:duration' # Duración en minutos
    REQUIRED_ATTENDEES_IPM_MEETING = 'urn:schemas:calendar:requiredattendees' # Asistentes obligatorios
    OPTIONAL_ATTENDEES_IPM_MEETING = 'urn:schemas:calendar:optionalattendees' # Asistentes opcionales
    RESOURCES_IPM_MEETING = 'urn:schemas:calendar:resources' # Recursos
    IMPORTANCE_IPM_MEETING = 'urn:schemas:calendar:importance' # Importancia de la cita o reunión (1=Alta, 2=Normal, 3=Baja)
    MEETING_STATUS_IPM_MEETING = 'urn:schemas:calendar:meetingstatus' # Estado de la reunión (0=No es reunión, 1=Reunión, 3=Reunión cancelada)
    CATEGORY_IPM_MEETING = 'urn:schemas:calendar:category' # Categoría de la cita o reunión
    BUSYSTATUS_IPM_MEETING = 'urn:schemas:calendar:busystatus' # Estado de disponibilidad (0=Libre, 1=Con reserva, 2=Ocupado, 3=Fuera de la oficina)
    ID_IPM_MEETING = 'urn:schemas:calendar:uid' # Identificador único del evento (UID)
    
    ############################################################################
    #####   Campos DASL para mensajes IPM.Task o IPM.ScheduleTask          #####
    ############################################################################
    SUBJECT_IPM_TASK = 'urn:schemas:task:subject' # Asunto de la tarea
    BODY_TEXT_IPM_TASK = 'urn:schemas:task:body' # Cuerpo de la tarea en texto plano
    OWNER_IPM_TASK = 'urn:schemas:task:owner' # Propietario de la tarea
    START_DATE_IPM_TASK = 'urn:schemas:task:startdate' # Fecha de inicio
    DUE_DATE_IPM_TASK = 'urn:schemas:task:duedate' # Fecha de vencimiento
    STATUS_IPM_TASK = 'urn:schemas:task:status' # Estado de la tarea (0=No iniciada, 1=En progreso, 2=Completada, 3=En espera, 4=Cancelada)
    PERCENT_COMPLETE_IPM_TASK = 'urn:schemas:task:percentcomplete' # Porcentaje de completitud (0-100)
    ACTUAL_WORK_IPM_TASK = 'urn:schemas:task:actualwork' # Trabajo real en minutos
    TOTAL_WORK_IPM_TASK = 'urn:schemas:task:totalwork' # Trabajo total en minutos
    ID_IPM_TASK = 'urn:schemas:task:uid' # Identificador único de la tarea
    
    ############################################################################
    #####   Campos DASL para mensajes IPM.Contact                          #####
    ############################################################################
    NAME_IPM_CONTACT = 'urn:schemas:contacts:fullname' # Nombre del contacto
    EMAIL_IPM_CONTACT = 'urn:schemas:contacts:email' # Email del contacto
    EMAIL2_IPM_CONTACT = 'urn:schemas:contacts:email2' # Segundo email del contacto
    EMAIL3_IPM_CONTACT = 'urn:schemas:contacts:email3' # Tercer email del contacto
    BUSINESS_PHONE_IPM_CONTACT = 'urn:schemas:contacts:businessphone' # Teléfono del trabajo del contacto
    MOBILEPHONE_IPM_CONTACT = 'urn:schemas:contacts:mobilephone' # Teléfono móvil del contacto
    HOME_PHONE_IPM_CONTACT = 'urn:schemas:contacts:homephone' # Teléfono de casa del contacto
    JOB_TITLE_IPM_CONTACT = 'urn:schemas:contacts:jobtitle' # Cargo del contacto
    ADDRESS_IPM_CONTACT = 'urn:schemas:contacts:address' # Dirección del contacto
    BIRTHDAY_IPM_CONTACT = 'urn:schemas:contacts:birthday' # Cumpleaños del contacto
    NOTES_IPM_CONTACT = 'urn:schemas:contacts:notes' # Notas del contacto
    BODY_IPM_CONTACT = 'urn:schemas:contacts:body' # Cuerpo en texto plano
    CATEGORY_IPM_CONTACT = 'urn:schemas:contacts:categories' # Categoría del contacto
    ID_IPM_CONTACT = 'urn:schemas:contacts:uid' # Identificador único del contacto
    
    ############################################################################
    #####   Campos DASL para mensajes IPM.StickyNote                       #####
    ############################################################################
    SUBJECT_IPM_STICKYNOTE = 'urn:schemas:note:subject' # Asunto de la nota
    BODY_TEXT_IPM_STICKYNOTE = 'urn:schemas:note:body' # Cuerpo de la nota en texto plano
    COLOR_IPM_STICKYNOTE = 'urn:schemas:note:color' #Color de la nota (0=Amarillo, 1=Azul, 2=Verde, 3=Rosa, 4=Naranja, 5=Morado, 6=Rojo)
    CATEGORY_IPM_STICKYNOTE = 'urn:schemas:note:categories' # Categoría de la nota
    CREATED_TIME_IPM_STICKYNOTE = 'urn:schemas:note:created'
    ID_IPM_STICKYNOTE = 'urn:schemas:note:uid' # Identificador único de la nota
    
    ############################################################################
    #####   Campos DASL para mensajes IPM.Post                             #####
    ############################################################################
    SUBJECT_IPM_POST = 'urn:schemas:post:subject' # Asunto del post
    BODY_TEXT_IPM_POST = 'urn:schemas:post:body' # Cuerpo del post en texto plano
    SENDER_EMAIL_IPM_POST = 'urn:schemas:post:fromemail' # Email del remitente
    SENDER_NAME_IPM_POST = 'urn:schemas:post:fromname' # Nombre del remitente
    IMPORTANCE_IPM_POST = 'urn:schemas:post:importance' # Importancia del post (1=Alta, 2=Normal, 3=Baja)
    HAS_ATTACHMENTS_IPM_POST = 'urn:schemas:post:hasattachment' # Si el post tiene adjuntos
    IS_READ_IPM_POST = 'urn:schemas:post:read' # Si el post ha sido leído
    CATEGORY_IPM_POST = 'urn:schemas:post:categories' # Categoría del post
    CREATED_TIME_IPM_POST = 'urn:schemas:post:created' # Fecha y hora de creación
    MODIFIED_TIME_IPM_POST = 'urn:schemas:post:modified' # Fecha y hora de modificación
    ID_IPM_POST = 'urn:schemas:post:messageid' # Identificador único del post
    
    ############################################################################
    #####   Campos DASL para mensajes IPM.Post                             #####
    ############################################################################
    SUBJECT_IPM_JOURNAL = 'urn:schemas:journal:subject' # Asunto de la entrada del diario
    BODY_TEXT_IPM_JOURNAL = 'urn:schemas:journal:body' # Cuerpo de la entrada del diario en texto plano
    FILENAME_IPM_JOURNAL = 'urn:schemas:journal:filename' # Nombre del archivo asociado
    TYPE_IPM_JOURNAL = 'urn:schemas:journal:type' # Tipo de entrada
    CATEGORY_IPM_JOURNAL = 'urn:schemas:journal:categories' # Categoría de la entrada
    CREATED_TIME_IPM_JOURNAL = 'urn:schemas:journal:created' # Fecha y hora de creación
    ID_IPM_JOURNAL = 'urn:schemas:journal:uid' # Identificador único de la entrada del diario
    


class OUTLOOKTYPERECIPENTS(int, Enum):
    TO = 1
    CC = 2
    BCC = 3
    

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
    received_after: Optional[datetime] = None  # ISO format date string
    received_before: Optional[datetime] = None  # ISO format date string
    conversation_topic: Optional[str] = None  # Corresponde al asunto del hilo de conversación (conversation topic)
    referenceid: Optional[list[str]] = None  # Corresponde al Message-ID de los emails referenciados (header References)
    msg_id: Optional[list[str]] = None  # Corresponde al Message-ID del email (header)
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
    store_folder: str
    standard_folder: Optional[OutlookStandarFolders] = OutlookStandarFolders.INBOX
    custom_folder: Optional[str] = None  # Esta es la ruta completa de la carpeta personalizada
    max_emails: Optional[int] = 500  # Número máximo de emails a obtener
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
    

class DataDownloadAttachments(BaseModel):
    email_ids: list[str]
    download_folder: str
    mark_as_read: Optional[bool] = False  # Marcar los emails obtenidos como leídos
    overwrite: Optional[bool] = True  # Sobrescribir archivos si ya existen
    only_filenames: Optional[list[str]] = None  # Si se especifica, solo se descargan los archivos con estos nombres
    only_extensions: Optional[list[str]] = None  # Si se especifica, solo se descargan los archivos con estas extensiones (sin el punto)
    ignore_extensions: Optional[list[str]] = None  # Si se especifica, no se descargan los archivos con estas extensiones (sin el punto)
    ignore_filenames: Optional[list[str]] = None  # Si se especifica, no se descargan los archivos con estos nombres
    create_subfolder_per_email: Optional[bool] = True  # Crear una subcarpeta por cada email, nombrada con el nombre y fecha de cada correo
    
    @field_validator('only_extensions', 'ignore_extensions', mode='before')
    def validate_extensions(cls, v):
        if v is None:
            return v
        if not isinstance(v, list):
            raise ValueError("Las extensiones deben ser una lista.")
        lista_sin_extension = [s.split('.')[-1] if not s.startswith('.') else s[1:] for s in v]
        return lista_sin_extension
    
    @field_validator('only_filenames', 'ignore_filenames', mode='before')
    def validate_filenames(cls, v):
        if v is None:
            return v
        if not isinstance(v, list):
            raise ValueError("Los nombres de archivo deben ser una lista.")
        
        filenames = [s.replace('.', '') if s.startswith('.') else s.split('.')[0] for s in v]
        return filenames
