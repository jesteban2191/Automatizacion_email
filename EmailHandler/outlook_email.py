from .email_interface import EmailInterface
from ..common import ConnectionInfo, DataGetEmails, DataFiltersEmails, QUERYDASL, IMPORTANCEEMAIL, LOGICOPERATOR, SUBJECTPREFIX, DataDownloadAttachments, DataSendEmail, EmailAttachmentInfo
import win32com.client
import pythoncom
import re
import pandas as pd
from time import time
from ..helpers import segundos_a_horas_minutos_segundos, remove_emojis, format_datetime, format_date_folder
from ..common import OUTLOOKTYPERECIPENTS, OutlookStandardFoldersstr, OutlookStandarFolders, DateTypes, OUTLOOKTYPEELEMENT, OUTLOOKTYPEATTACHMENTS
import os
from pathlib import Path
from typing import Optional
from datetime import datetime

class OutlookEmail(EmailInterface):

    def __init__(self, conn: win32com.client.CDispatch, app: win32com.client.CDispatch):
        self._conn = conn # Este es el objeto Namespace autenticado
        self._app = app
        self.tiempo_transformacion_datos_acumulado = 0
        self.tiempo_descarga_acumulado = 0
        self.authenticated_email = self.get_main_mailbox_name()
        print(self.authenticated_email)
        pass
        
    
    def get_emails(self, datagetemails: DataGetEmails):
        
        
        start=time()
        #########################################################################################
        #####          Obtengo todos los parámetros para descagar los emails            #########
        #########################################################################################
        self.store_folder = datagetemails.store_folder_mail if datagetemails.store_folder_mail else self.authenticated_email
        self.standard_folders = datagetemails.standard_folder_mail
        self.custom_folder = datagetemails.custom_folder_mail
        self.max_emails = datagetemails.max_emails
        self.mark_as_read = datagetemails.mark_as_read
        self.page_next = datagetemails.page_next
        self.download_attachments = datagetemails.download_attachments
        self.attachments_settings = datagetemails.attachments_settings if datagetemails.attachments_settings else None
        self.tiempo_descarga = 0
        self.tiempo_transformacion_datos = 0
        self.tiempo_transformacion_datos_acumulado = 0 if self.page_next is None else self.tiempo_transformacion_datos_acumulado
        self.tiempo_descarga_acumulado = 0 if self.page_next is None else self.tiempo_descarga_acumulado
        
        
        
        #########################################################################################
        #####               Creo el query para obtener los emails                       #########
        #########################################################################################
        
        self.datafilter = datagetemails.filters
        if self.datafilter:
            self.query = self.create_query(self.datafilter)
        else:
            self.query = None
            
        print(f"Query creado: {self.query}")
        
        
        #########################################################################################
        #####      Preparo la configuración de descarga de adjuntos si es necesario     #########
        #########################################################################################
        if self.download_attachments:
            
            self.path_download_folder = Path(self.attachments_settings.download_folder)
            self.path_download_folder.mkdir(parents=True, exist_ok=True)
            self.overwrite = self.attachments_settings.overwrite
            self.only_filenames = self.attachments_settings.only_filenames or []
            self.only_extensions = self.attachments_settings.only_extensions or []
            self.ignore_extensions = self.attachments_settings.ignore_extensions or []
            self.ignore_filenames = self.attachments_settings.ignore_filenames or []
            self.create_subfolder_per_email = self.attachments_settings.create_subfolder_per_email
            self.subfolder_name = self.attachments_settings.name_subfolder_per_email if self.attachments_settings.name_subfolder_per_email else "{index}_{subject}_{receiveddate}"

        #########################################################################################
        #####               Obtengo los emails según los parámetros                     #########
        #########################################################################################

        folder = self.validate_folder(self.standard_folders, self.custom_folder)
        
        
        if self.query:
            
            filtered_messages = folder.Items.Restrict(self.query)
        else:
            filtered_messages = folder.Items
            
        
        filtered_messages.Sort("ReceivedTime", True)  # Ordenar por fecha de recepción descendente
        
        self.tiempo_transformacion_datos = time() - start
        tiempo_transformacion_datos_acumulado = self.tiempo_transformacion_datos + self.tiempo_transformacion_datos_acumulado
        tiempo_transformacion_datos = segundos_a_horas_minutos_segundos(tiempo_transformacion_datos_acumulado)

        self.total_emails = filtered_messages.Count
        self.page_number = 1 if self.page_next is None else self.page_next
        self.page_size = self.max_emails
        self.start = (self.page_number - 1) * self.page_size
        self.end = self.start + self.page_size
        filtered_messages = list(filtered_messages)  # Convierte a lista para poder hacer slicing
        filtered_messages = filtered_messages[self.start:self.end]        
        self.page_next_result = self.page_number + 1 if self.end < self.total_emails else None
        self.total_adjuntos = self.count_att_filtered(filtered_messages, self.only_extensions, self.only_filenames, self.ignore_extensions, self.ignore_filenames) if self.download_attachments else 0

        emails = []
        
        start=time()
        
        for i, message in enumerate(filtered_messages, start=self.start):
            message_class = getattr(message, 'MessageClass', None)
            is_mail = message_class == 'IPM.Note'
            is_appointment = message_class in ['IPM.Appointment', 'IPM.Schedule.Meeting.Request']
            
            attachments_files = self.get_list_of_attachments_filtered(message, self.only_extensions, self.only_filenames, self.ignore_extensions, self.ignore_filenames) if self.download_attachments else []
            

            if is_mail:
                email_data = {
                    'Tipo': 'Correo',
                    'Subject': remove_emojis(getattr(message, 'Subject', None)),
                    'Sender': f"{self.get_sender_str(message)}",
                    'To': f"{self.get_recipients_str(message, OUTLOOKTYPERECIPENTS.TO.value)}",
                    'CC': f"{self.get_recipients_str(message, OUTLOOKTYPERECIPENTS.CC.value)}",
                    'BCC': f"{self.get_recipients_str(message, OUTLOOKTYPERECIPENTS.BCC.value)}",
                    'ReceivedTime': format_datetime(getattr(message, 'ReceivedTime', None)),
                    'SentOn': format_datetime(getattr(message, 'SentOn', None)),
                    'Body': remove_emojis(getattr(message, 'Body', None)),
                    'HTMLBody': remove_emojis(getattr(message, 'HTMLBody', None)),
                    'Num_of_Attachments': getattr(getattr(message, 'Attachments', None), 'Count', 0),
                    'IsRead': getattr(message, 'UnRead', None) == 0,
                    'Importance': "LOW" if getattr(message, 'Importance', None) == 0 else "NORMAL" if getattr(message, 'Importance', None) == 1 else "HIGH" if getattr(message, 'Importance', None) == 2 else "Unknown",
                    'ConversationTopic': remove_emojis(getattr(message, 'ConversationTopic', None)),
                    'ConversationID': getattr(message, 'ConversationID', None),
                    'MessageID': getattr(message, 'EntryID', None),
                    'MessageClass': message_class,
                    'Categories': getattr(message, 'Categories', None),
                    'Size': getattr(message, 'Size', None),
                    'Attachments': [(getattr(att, 'FileName', None), getattr(att, 'Index', None)) for att in getattr(message, 'Attachments', [])] if getattr(message, 'Attachments', None) and hasattr(getattr(message, 'Attachments', None), '__iter__') else []
                }
            elif is_appointment:
                email_data = {
                    'Tipo': 'Cita/Reunión',
                    'Subject': getattr(message, 'Subject', None),
                    'Organizer': f"{self.get_meeting_organizer(message)}",
                    'Start': self.get_meeting_start(message),
                    'End': self.get_meeting_end(message),
                    'Body': getattr(message, 'Body', None),
                    'HTMLBody': getattr(message, 'HTMLBody', None),
                    'Recipients': f"{self.get_recipients_str(message, None)}",  # O ajusta para todos los tipos si lo deseas
                    'MessageClass': message_class,
                    'Categories': getattr(message, 'Categories', None),
                    'Attachments': [(getattr(att, 'FileName', None), getattr(att, 'Index', None)) for att in getattr(message, 'Attachments', [])] if getattr(message, 'Attachments', None) and hasattr(getattr(message, 'Attachments', None), '__iter__') else []
                }
            else:
                email_data = {
                    'Tipo': f'Otro ({message_class})',
                    'Subject': getattr(message, 'Subject', None),
                    'MessageClass': message_class,
                }
                
            if self.download_attachments:
                #########################################################################################
                #####               Aquí voy a descargar los adjuntos que correspondan          #########
                #########################################################################################
                folder_path = self.create_folder_to_download_attachments(message, self.path_download_folder, self.subfolder_name, email_data, self.create_subfolder_per_email, i) if attachments_files else None

                file_paths = [folder_path / file_name for file_name in attachments_files] if folder_path else []

                if file_paths:
                    folder_path.mkdir(parents=True, exist_ok=True)
                else:
                    folder_path = None
                
                if i == 0:
                    total_att_downloaded = 0
                    current_att_downloaded = 0
                    
                if file_paths:
                    #########################################################################################
                    #####               Sí hay archivos que descargar los descargo               #########
                    #########################################################################################
                    for j, dest_path in enumerate(file_paths):
                        if dest_path.exists() and not self.overwrite:
                            #print(f"El archivo {dest_path} ya existe y overwrite está establecido en False. Saltando descarga.")
                            continue
                        elif dest_path.exists() and self.overwrite:
                            #print(f"El archivo {dest_path} ya existe pero overwrite está establecido en True. Sobrescribiendo archivo.")
                            dest_path.unlink()
                            
                        try:
                            attachment = next((att for att in getattr(message, 'Attachments', []) if getattr(att, 'FileName', None) == dest_path.name), None)
                            if attachment:
                                attachment.SaveAsFile(str(dest_path))
                                self.tiempo_descarga_adjuntos = time() - start
                        except Exception as e:
                            raise Exception(f"Error al guardar el adjunto {dest_path}: {e}")
                        
                        current_att_downloaded = 1 + j
                        total_att_downloaded += current_att_downloaded
                        os.system('cls' if os.name == 'nt' else 'clear')
                        self.tiempo_descarga = time() - start
                        tiempo_descarga_acumulado = self.tiempo_descarga_acumulado + self.tiempo_descarga
                        print(f"""
                                ---------------------------------------------------------------------------------------------------------
                                    Total Emails to download --> {self.total_emails}, current_page --> {self.page_number}, page_size --> {self.page_size}
                                    Total Attachments --> {self.total_adjuntos}
                                    folder_name --> {folder_path if folder_path else "No folder created"}
                                    Email downloaded --> {i + 1} of {self.total_emails} ({round(((i + 1)/self.total_emails)*100,2)}%)
                                    Attachments downloaded --> {total_att_downloaded} of {self.total_adjuntos} ({round(((total_att_downloaded)/self.total_adjuntos)*100,2)}%)
                                    Tiempo en tratamiento de datos --> {tiempo_transformacion_datos}
                                    Tiempo en descarga de adjuntos --> {segundos_a_horas_minutos_segundos(self.tiempo_descarga_acumulado)}
                                ---------------------------------------------------------------------------------------------------------""")
                else:
                    #########################################################################################
                    #####    Si no tengo adjuntos que descargar en el correo actual, continuo       #########
                    #########################################################################################
                    os.system('cls' if os.name == 'nt' else 'clear')
                    self.tiempo_descarga = time() - start
                    tiempo_descarga_acumulado = self.tiempo_descarga_acumulado + self.tiempo_descarga
                    print(f"""
                            ---------------------------------------------------------------------------------------------------------
                                Total Emails to download --> {self.total_emails}, current_page --> {self.page_number}, page_size --> {self.page_size}
                                Total Attachments --> {self.total_adjuntos}
                                folder_name --> {folder_path if folder_path else "No folder created"}
                                Email downloaded --> {i + 1} of {self.total_emails} ({round(((i + 1)/self.total_emails)*100,2)}%)
                                Attachments downloaded --> {total_att_downloaded} of {self.total_adjuntos} ({round(((total_att_downloaded)/self.total_adjuntos)*100,2)}%)
                                Tiempo en tratamiento de datos --> {tiempo_transformacion_datos}
                                Tiempo en descarga de adjuntos --> {segundos_a_horas_minutos_segundos(self.tiempo_descarga_acumulado)}
                            ---------------------------------------------------------------------------------------------------------""")
                
                email_data['Attachment_Folder'] = str(folder_path) if folder_path else "No folder created"
                
            else:
                #########################################################################################
                #####    En caso de no tener que descargar adjuntos                             #########
                #########################################################################################
                self.tiempo_descarga = time() - start
                tiempo_descarga_acumulado = self.tiempo_descarga_acumulado + self.tiempo_descarga
                os.system('cls')
                print(f"""
                    -------------------------------------------------------------------------------------------------
                            Total Emails to download --> {self.total_emails}, current_page --> {self.page_number}, page_size --> {self.page_size}
                            Email downloaded --> {i + 1} of {self.total_emails} ({round(((i + 1)/self.total_emails)*100,2)}%)
                            Tiempo en tratamiento de datos --> {tiempo_transformacion_datos}
                            Tiempo en descarga de datos --> {segundos_a_horas_minutos_segundos(tiempo_descarga_acumulado)}
                    --------------------------------------------------------------------------------------------------""")
            
            email_data['Attachment_Folder'] = str(folder_path) if folder_path else "No folder created"
            emails.append(email_data)
            
        df_emails = pd.DataFrame(emails)
        
        self.tiempo_transformacion_datos_acumulado = 0 if self.page_next_result is None else tiempo_transformacion_datos_acumulado
        self.tiempo_descarga_acumulado = 0 if self.page_next_result is None else tiempo_descarga_acumulado
        
        result = {
            "data": df_emails.astype(str),
            "page_number": self.page_number,
            "page_size": self.page_size,
            "total_emails": self.total_emails,
            "has_more": self.end < self.total_emails,
            "page_next": self.page_next_result
        }

        return result
    
    
    
    def get_main_mailbox_name(self):
        # Obtiene el correo SMTP del usuario autenticado
        try:
            smtp_address = self._conn.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        except Exception:
            smtp_address = self._conn.CurrentUser.Address
        # Busca el store cuyo nombre coincide con el correo SMTP
        for i in range(1, self._conn.Folders.Count + 1):
            store = self._conn.Folders[i]
            if store.Name == smtp_address:
                return store.Name
        # Si no lo encuentra, retorna el primero que no sea Favoritos/Public Folders/Archivar en línea
        for i in range(1, self._conn.Folders.Count + 1):
            store = self._conn.Folders[i]
            if store.Name not in ["Favoritos", "Public Folders", "Archivar en línea"]:
                return store.Name
        return None
    
    #################################################################################################################################
    #####                               Función para validar el folder solicitado                                           #########
    #################################################################################################################################
    def validate_folder(self, standard_folder: OutlookStandarFolders, custom_folder: Optional[str] = None) -> win32com.client.CDispatch:
        
        base_path = f"{self.store_folder}\\{OutlookStandardFoldersstr[standard_folder.name].value}"
        custom_folder_outlook_path = custom_folder.replace("/", "\\") if custom_folder else None
        full_folder = f"{base_path}\\{custom_folder_outlook_path}" if custom_folder else f"{base_path}"
        list_path_folders = self.get_path_folders()
        if full_folder not in list_path_folders:
            raise ValueError(f"La ruta especificada '{full_folder}' no existe en Outlook. Rutas disponibles: {list_path_folders}")
                
        if custom_folder:
            custom_folder = custom_folder.split('/')
        
        folder = self._conn.GetDefaultFolder(standard_folder.value)
        if custom_folder:
            for subfolder in custom_folder:
                try:
                    folder = folder.Folders[subfolder]
                except Exception:
                    raise ValueError(f"La subcarpeta '{subfolder}' no existe en '{folder.Name}'.")
                
        return folder
        
    
    
    #################################################################################################################################
    #####                               Función para creación de consulta SQL DASL                                          #########
    #################################################################################################################################
    def create_query (self, datafilters: DataFiltersEmails) -> str:
        query_header = "@SQL= ("
        self.subject = datafilters.subject
        self.sender_email = datafilters.sender_email
        self.sender = datafilters.sender
        self.recipient_email = datafilters.recipient_email
        self.recipient = datafilters.recipient
        self.cc_email = datafilters.cc_email
        self.bcc_email = datafilters.bcc_email
        self.cc = datafilters.cc
        self.bcc = datafilters.bcc
        self.body = datafilters.body
        self.has_attachments = datafilters.has_attachments
        self.is_read = datafilters.is_read
        self.received_after = datafilters.received_after
        self.received_before = datafilters.received_before
        self.conversation_topic = datafilters.conversation_topic
        self.referenceid = datafilters.referenceid
        self.msg_id = datafilters.msg_id
        self.importance_email = datafilters.importance_email
        self.logic_operator = datafilters.logic_operator
        self.subject_prefix = datafilters.subject_prefix
        self.logic_operator_between_senders = datafilters.logic_operator_between_senders
        self.logic_operator_between_recipients = datafilters.logic_operator_between_recipients
        
        query_parts = []
        
        if self.subject:
            if "%" in self.subject:
                query_parts.append(f"""({QUERYDASL.SUBJECT_IPM_NOTE.value} LIKE '{self.subject.lower()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_IPM_NOTE.value} LIKE '{self.subject.upper()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_IPM_NOTE.value} LIKE '{self.subject.title()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_IPM_NOTE.value} LIKE '{self.subject.capitalize()}'
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_IPM_NOTE.value} LIKE '{self.subject}')""")
            else:
                query_parts.append(f"""({QUERYDASL.SUBJECT_IPM_NOTE.value} = '{self.subject.lower()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_IPM_NOTE.value} = '{self.subject.upper()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_IPM_NOTE.value} = '{self.subject.title()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_IPM_NOTE.value} = '{self.subject.capitalize()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_IPM_NOTE.value} = '{self.subject}')""")

        if self.sender_email:
            senders = [f"""{QUERYDASL.SENDER_EMAIL_IPM_NOTE.value} = '{email.lower()}'""" for email in self.sender_email]
            query_parts.append(f"(" + f" {self.logic_operator_between_senders.value} ".join(senders) + ")")

        if self.recipient_email:
            recipients = [f"{QUERYDASL.RECIPIENT_EMAIL.value} = '{email.lower()}'" for email in self.recipient_email]
            query_parts.append(f"(" + f" {self.logic_operator_between_recipients.value} ".join(recipients) + ")")
            
        if self.sender:
            senders = [f"""({QUERYDASL.SENDER_NAME.value} LIKE '%{name.lower()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.SENDER_NAME.value} LIKE '%{name.title()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.SENDER_NAME.value} LIKE '%{name.capitalize()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.SENDER_NAME.value} LIKE '%{name.upper()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.SENDER_NAME.value} LIKE '%{name}%')""" for name in self.sender]
            query_parts.append(f"(" + f" {self.logic_operator_between_senders.value} ".join(senders) + ")")
        
        if self.recipient:
            recipients = [f"""({QUERYDASL.RECIPIENT_NAME.value} LIKE '%{name.lower()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.RECIPIENT_NAME.value} LIKE '%{name.title()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.RECIPIENT_NAME.value} LIKE '%{name.capitalize()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.RECIPIENT_NAME.value} LIKE '%{name.upper()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.RECIPIENT_NAME.value} LIKE '%{name}%')""" for name in self.recipient]
            query_parts.append(f"(" + f" {self.logic_operator_between_recipients.value} ".join(recipients) + ")")

        if self.cc_email:
            cc_recipients = [f"{QUERYDASL.CC_EMAIL.value} = '{email.lower()}'" for email in self.cc_email]
            query_parts.append(f"(" + f" {self.logic_operator_between_recipients.value} ".join(cc_recipients) + ")")

        if self.bcc_email:
            bcc_recipients = [f"{QUERYDASL.BCC_EMAIL.value} = '{email.lower()}'" for email in self.bcc_email]
            query_parts.append(f"(" + f" {self.logic_operator_between_recipients.value} ".join(bcc_recipients) + ")")

        if self.cc:
            cc_names = [f"""({QUERYDASL.CC_NAME.value} LIKE '%{name.lower()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.CC_NAME.value} LIKE '%{name.title()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.CC_NAME.value} LIKE '%{name.capitalize()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.CC_NAME.value} LIKE '%{name.upper()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.CC_NAME.value} LIKE '%{name}%')""" for name in self.cc]
            query_parts.append(f"(" + f" {self.logic_operator_between_recipients.value} ".join(cc_names) + ")")

        if self.bcc:
            bcc_names = [f"""({QUERYDASL.BCC_NAME.value} LIKE '%{name.lower()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.BCC_NAME.value} LIKE '%{name.title()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.BCC_NAME.value} LIKE '%{name.capitalize()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.BCC_NAME.value} LIKE '%{name.upper()}%' {LOGICOPERATOR.OR.value}
                       {QUERYDASL.BCC_NAME.value} LIKE '%{name}%')""" for name in self.bcc]
            query_parts.append(f"(" + f" {self.logic_operator_between_recipients.value} ".join(bcc_names) + ")")

        if self.body:
            if "%" in self.body:
                body_value = self.body.lower()
                query_parts.append(f"""({QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value}' {LOGICOPERATOR.OR.value} 
                                   {QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value.upper()}' {LOGICOPERATOR.OR.value} 
                                   {QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value.title()}' {LOGICOPERATOR.OR.value} 
                                   {QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value.capitalize()}' {LOGICOPERATOR.OR.value} 
                                   {QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value}')""")
            else:
                body_value = f"%{self.body.lower()}%"
                query_parts.append(f"""({QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value}' {LOGICOPERATOR.OR.value} 
                                   {QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value.upper()}' {LOGICOPERATOR.OR.value} 
                                   {QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value.title()}' {LOGICOPERATOR.OR.value} 
                                   {QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value.capitalize()}' {LOGICOPERATOR.OR.value} 
                                   {QUERYDASL.BODY_TEXT_IPM_NOTE.value} LIKE '{body_value}')""")

        if self.has_attachments is not None:
            has_attachments_value = 1 if self.has_attachments else 0
            query_parts.append(f"({QUERYDASL.HAS_ATTACHMENTS_IPM_NOTE.value} = {has_attachments_value})")
            
        if self.is_read is not None:
            is_read_value = 1 if self.is_read else 0
            query_parts.append(f"({QUERYDASL.IS_READ_IPM_NOTE.value} = {is_read_value})")

        if self.received_after and self.received_before:
            if self.received_after > self.received_before:
                raise ValueError("received_after no puede ser mayor que received_before")
            date_filter = (
                f"({QUERYDASL.RECEIVED_TIME_IPM_NOTE.value} >= '{self.received_after.strftime('%m/%d/%Y %H:%M:%S')}' "
                f"AND {QUERYDASL.RECEIVED_TIME_IPM_NOTE.value} <= '{self.received_before.strftime('%m/%d/%Y %H:%M:%S')}')"
            )
        elif self.received_after:
            date_filter = (
                f"({QUERYDASL.RECEIVED_TIME_IPM_NOTE.value} >= '{self.received_after.strftime('%m/%d/%Y %H:%M:%S')}')"
            )
        elif self.received_before:
            date_filter = (
                f"({QUERYDASL.RECEIVED_TIME_IPM_NOTE.value} <= '{self.received_before.strftime('%m/%d/%Y %H:%M:%S')}')"
            )
        else:
            date_filter = None

        if date_filter:
            query_parts.append(date_filter)
            
        if self.conversation_topic:
            if "%" in self.conversation_topic:
                query_parts.append(f"""({QUERYDASL.CONVERSATION_TOPIC.value} LIKE '{self.conversation_topic.lower()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.CONVERSATION_TOPIC.value} LIKE '{self.conversation_topic.upper()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.CONVERSATION_TOPIC.value} LIKE '{self.conversation_topic.title()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.CONVERSATION_TOPIC.value} LIKE '{self.conversation_topic.capitalize()}'
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.CONVERSATION_TOPIC.value} LIKE '%{self.conversation_topic}%')""")
            else:
                query_parts.append(f"""({QUERYDASL.CONVERSATION_TOPIC.value} = '{self.conversation_topic.lower()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.CONVERSATION_TOPIC.value} = '{self.conversation_topic.upper()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.CONVERSATION_TOPIC.value} = '{self.conversation_topic.title()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.CONVERSATION_TOPIC.value} = '{self.conversation_topic.capitalize()}'
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.CONVERSATION_TOPIC.value} = '{self.conversation_topic}')""")

        if self.referenceid:
            references = [f"{QUERYDASL.REFERENCE_ID.value} = '{ref}'" for ref in self.referenceid]
            query_parts.append(f"(" + " OR ".join(references) + "))")
        if self.msg_id:
            msgids = [f"{QUERYDASL.ID_IPM_NOTE.value} = '{ref}'" for ref in self.msg_id]
            query_parts.append(f"(" + " OR ".join(msgids) + "))")

        if self.importance_email:
            query_parts.append(f"({QUERYDASL.IMPORTANCE_IPM_NOTE.value} = {self.importance_email.value})")
            
        if self.subject_prefix:
            query_parts.append(f"""({QUERYDASL.SUBJECT_PREFIX_IPM_NOTE.value} = '{self.subject_prefix.value.lower()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_PREFIX_IPM_NOTE.value} = '{self.subject_prefix.value.upper()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_PREFIX_IPM_NOTE.value} = '{self.subject_prefix.value.title()}' 
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_PREFIX_IPM_NOTE.value} = '{self.subject_prefix.value.capitalize()}'
                                   {LOGICOPERATOR.OR.value} {QUERYDASL.SUBJECT_PREFIX_IPM_NOTE.value} = '{self.subject_prefix.value}')""")

        if not query_parts:
            query = None
        else:
            query = query_header + f" {self.logic_operator.value} ".join(query_parts) + ")"
            
        return query
        
    
    #################################################################################################################################
    #####                               Función para obtener el correo SMTP del remitente                                   #########
    #################################################################################################################################         
    def get_sender_smtp(self, message) -> str:
        sender = getattr(message, 'Sender', None)
        if sender is not None:
            # Si es un usuario de Exchange, intenta obtener el correo SMTP real
            if getattr(message, 'SenderEmailType', None) == 'EX':
                exchange_user = getattr(sender, 'GetExchangeUser', lambda: None)()
                if exchange_user is not None:
                    return getattr(exchange_user, 'PrimarySmtpAddress', None)
            # Si no es Exchange, usa el campo estándar
            return getattr(message, 'SenderEmailAddress', None)
        return None
    
    
    #################################################################################################################################
    #####                               Función para obtener el nombre del remitente                                        #########
    #################################################################################################################################
    def get_sender_str(self, message) -> str:
        sender_name = getattr(message, 'SenderName', '') or ''
        sender_email = self.get_sender_smtp(message) or ''
        if sender_name and sender_email:
            sender_str = f"{sender_name} ({sender_email})"
        elif sender_name:
            sender_str = sender_name
        elif sender_email:
            sender_str = sender_email
        else:
            sender_str = ''
            
        return sender_str
    
        
    #################################################################################################################################
    #####                               Función para obtener el correo SMTP de quien recibe                                 #########
    #################################################################################################################################    
    def get_recipient_smtp(self, recipient) -> str:
        address_entry = getattr(recipient, 'AddressEntry', None)
        if address_entry is not None:
            if getattr(address_entry, 'Type', None) == 'EX':
                exchange_user = getattr(address_entry, 'GetExchangeUser', lambda: None)()
                if exchange_user is not None:
                    return getattr(exchange_user, 'PrimarySmtpAddress', None)
            return getattr(recipient, 'Address', None)
        return None
    
    
    #################################################################################################################################
    #####                               Función para obtener el nombre de quien recibe                                      #########
    #################################################################################################################################
    def get_recipients_str(self, message, recipient_type) -> str:
        recipients = getattr(message, 'Recipients', [])
        if recipient_type is not None:
            recipients = [r for r in recipients if hasattr(r, 'Type') and r.Type == recipient_type]
        recipient_strs = []
        for r in recipients:
            name = getattr(r, 'Name', '') or ''
            email = self.get_recipient_smtp(r) or ''
            if name and email:
                recipient_strs.append(f"{name} ({email})")
            elif name:
                recipient_strs.append(name)
            elif email:
                recipient_strs.append(email)
        return "; ".join(recipient_strs)
    
    
    #################################################################################################################################
    #####                               Función para obtener el correo SMTP del organizador                                 #########
    #################################################################################################################################    
    def get_organizer_smtp(self, message) -> str:
        organizer = getattr(message, 'Organizer', None)
        if not organizer:
            return ''
        # Buscar en Recipients el que coincida con el nombre del Organizer
        for r in getattr(message, 'Recipients', []):
            if getattr(r, 'Name', '') == organizer:
                email = self.get_recipient_smtp(r)
                if email:
                    return f"{organizer} ({email})"
                else:
                    return organizer
        # Si no se encuentra, devolver solo el nombre
        return organizer
    
    
    #################################################################################################################################
    #####                               Función para obtener el nombre del organizador                                      #########
    #################################################################################################################################
    def get_meeting_organizer(self, message) -> str:
        organizer = getattr(message, 'Organizer', None)
        if organizer:
            return organizer
        sender = self.get_sender_str(message)
        if sender:
            return sender
        # Opcional: buscar en Recipients si coincide el nombre
        return ''
    
    
    #################################################################################################################################
    #####                          Función para obtener la fecha y hora de inicio de la solicitud de reunión                #########
    #################################################################################################################################
    def get_meeting_start(self, message):
        for attr in ['Start', 'StartUTC', 'MeetingStartTime', 'OriginalStart', 'AppointmentStart']:
            value = getattr(message, attr, None)
            if value:
                return value
        if getattr(message, 'MessageClass', None) == 'IPM.Schedule.Meeting.Request':
            appointment = getattr(message, 'GetAssociatedAppointment', lambda x=None: None)(False)
            if appointment:
                start = getattr(appointment, 'Start', None)
                if start:
                    return format_datetime(start)
        return None
    
    
    #################################################################################################################################
    #####                          Función para obtener la fecha y hora de fin de la solicitud de reunión                   #########
    #################################################################################################################################    
    def get_meeting_end(self, message):
        for attr in ['End', 'EndUTC', 'MeetingEndTime', 'OriginalEnd', 'AppointmentEnd']:
            value = getattr(message, attr, None)
            if value:
                return value
        # Si no hay End pero sí Start y Duration
        start = self.get_meeting_start(message)
        duration = getattr(message, 'Duration', None)
        if start and duration:
            try:
                # Duration está en minutos
                return start + datetime.timedelta(minutes=duration)
            except Exception:
                pass
        
        if getattr(message, 'MessageClass', None) == 'IPM.Schedule.Meeting.Request':
            appointment = getattr(message, 'GetAssociatedAppointment', lambda x=None: None)(False)
            if appointment:
                End = getattr(appointment, 'End', None)
                if End:
                    return format_datetime(End)

        return None
    
    
    #################################################################################################################################
    #####                          Función para obtener la ruta de las carpetas de Outlook                                  #########
    #################################################################################################################################
    def get_path_folders(self) -> list:
        list_paths = []
        stores = self._conn.Folders
        stores_names = [store.Name for store in stores]
        print(f"Stores disponibles: {stores_names}")
        store_name = [store for store in stores_names if store == self.store_folder]
        store = next((s for s in stores if s.Name == self.store_folder), None)
        if store: 
            list_paths = [f"{self.store_folder}\\{folder.Name}" for folder in store.Folders]
            print(f"Folders disponibles en: {list_paths}")
        # for i in range(1, self._conn.Folders.Count + 1):
        #     store = self._conn.Folders[i]
        #     print(f"Store: {store.Name} index {i}")
        #     if self.store_folder == store.Name:
                
        #         for j in range(1, store.Folders.Count + 1):
        #             try:
        #                 folder = store.Folders[j]
        #                 path = f"{store.Name}\\{folder.Name}"
        #                 list_paths.append(path)
        #             except Exception as e:
        #                 print(f"Error accediendo a folder index {j}: {e}")
        #             #print(f"  Carpeta: {folder.Name} - Path: {path}")
            
        return list_paths
    
    
    #################################################################################################################################
    #####                          Función obtener los adjuntos que cumplen con los requisitos                              #########
    #################################################################################################################################
    def get_list_of_attachments_filtered (self, messages, only_extensions, only_filenames, ignore_extensions, ignore_filenames) -> list:
        
        attachments_files = []
        attachments_filenames = [getattr(att, 'FileName', None) for att in getattr(messages, 'Attachments', []) if getattr(att, 'FileName', None) is not None]
        
        attachments_files = [
            att for att in attachments_filenames
            if att is not None
            and not any(att.endswith(ext) for ext in ignore_extensions)
            and att not in ignore_filenames
        ]

        # Si hay filtros de inclusión, aplicar
        if only_extensions or only_filenames:
            attachments_files = [
                att for att in attachments_files
                if (any(att.endswith(ext) for ext in only_extensions) 
                    or att in only_filenames)
            ]
            
        return attachments_files

    #################################################################################################################################
    #####                   Función para crear la carpeta donde se van a guardar los adjuntos descargados                   #########
    #################################################################################################################################
    def create_folder_to_download_attachments (self, messages, path_download_folder,subfolder_name, email_data, create_subfolder_per_email, index) -> Path:

            if "{subject}" in subfolder_name:
                subfolder_name = subfolder_name.replace("{subject}", email_data.get('Subject', 'No_Subject').strip() or 'No_Subject')
            if "{recivedtime}" in subfolder_name:
                subfolder_name = subfolder_name.replace("{recivedtime}", format_date_folder(getattr(messages, 'ReceivedTime', datetime.now()), DateTypes.DATETIME.value))
            if "{reciveddate}" in subfolder_name:
                subfolder_name = subfolder_name.replace("{reciveddate}", format_date_folder(getattr(messages, 'ReceivedTime', datetime.now()), DateTypes.DATE.value))
            if "{sender_mail}" in subfolder_name:
                subfolder_name = subfolder_name.replace("{sender_mail}", self.get_sender_str(messages).strip() or 'No_Sender')
            if "{index}" in subfolder_name:
                subfolder_name = subfolder_name.replace("{index}", str(index + 1))

            email_folder = f"{subfolder_name}".strip() or f"{index + 1}_{email_data.get('Subject', 'No_Subject').strip() or 'No_Subject'}_{format_date_folder(getattr(messages, 'ReceivedTime', datetime.now()), DateTypes.DATETIME.value)}"

            folder_path = path_download_folder / (email_folder if create_subfolder_per_email else "")
            
            return folder_path
            
            
            
    
    #################################################################################################################################
    #####                          Función para contar los adjuntos a descargar                                             #########
    #################################################################################################################################
    def count_att_filtered (self, filtered_messages, only_extensions, only_filenames, ignore_extensions, ignore_filenames) -> int:
        total_attachments = 0
        for msg in filtered_messages:
            attachments_filenames = [getattr(att, 'FileName', None) for att in getattr(msg, 'Attachments', []) if getattr(att, 'FileName', None) is not None]
            attachments_files = [
                att for att in attachments_filenames
                if att is not None
                and not any(att.endswith(ext) for ext in ignore_extensions)
                and att not in ignore_filenames
            ]
            if only_extensions or only_filenames:
                attachments_files = [
                    att for att in attachments_files
                    if (any(att.endswith(ext) for ext in self.only_extensions) 
                        or att in self.only_filenames)
                ]
            total_attachments += len(attachments_files)
        return total_attachments





    #################################################################################################################################
    #####                          Función para enviar correos con o sin adjuntos                                           #########
    #################################################################################################################################
    def send_email(self, datasentemail: DataSendEmail):
        #########################################################################################
        #####     Obtengo todos los parámetros para enviar los emails con o sin adjuntos #########
        #########################################################################################
        self.subject = datasentemail.subject if datasentemail.subject else "No Subject"
        self.body = datasentemail.body if datasentemail.body else ""
        self.to_recipients = datasentemail.to_recipients_email if datasentemail.to_recipients_email else []
        self.cc_recipients = datasentemail.cc_recipients_email if datasentemail.cc_recipients_email else []
        self.bcc_recipients = datasentemail.bcc_recipients_email if datasentemail.bcc_recipients_email else []
        self.importance_email = datasentemail.importance_email if datasentemail.importance_email else IMPORTANCEEMAIL.NORMAL
        self.attachments = datasentemail.attachments if datasentemail.attachments else []
        self.is_html = datasentemail.is_html
        self.body = datasentemail.body if datasentemail.body else ""
        self.read_receipt = datasentemail.read_receipt
        self.delivery_receipt = datasentemail.delivery_receipt # Solicitar acuse de entrega
        self.send_on_behalf = datasentemail.send_on_behalf # Enviar en nombre de otro usuario (debe tener permisos)
        self.save_copy_sent_items = datasentemail.save_copy_sent_items # Guardar una copia en la carpeta de elementos enviados
        self.connection_info = datasentemail.connection_info

        #########################################################################################
        #####               Creo el objeto del email nuevo                              #########
        #########################################################################################
        email = self._app.CreateItem(OUTLOOKTYPEELEMENT.MAIL.value)  # 0 indica un nuevo correo
        
        #########################################################################################
        #####               Empiezo a armar las partes del correo nuevo                 #########
        #########################################################################################
        email.subject = self.subject
        email.body = self.body if not self.is_html else ""
        email.HTMLBody = self.body if self.is_html else ""
        email.importance = self.importance_email.value  # 1=Alta, 2=Normal, 3=Baja
        email.To = "; ".join(self.to_recipients) if self.to_recipients else ""
        email.CC = "; ".join(self.cc_recipients) if self.cc_recipients else ""
        email.BCC = "; ".join(self.bcc_recipients) if self.bcc_recipients else ""
        if self.read_receipt:
            email.ReadReceiptRequested = True
        if self.delivery_receipt:
            email.DeliveryReceiptRequested = True
            
        #########################################################################################
        #####       Verifico las rutas de los adjuntos y agrego los adjuntos al email      ######
        #########################################################################################
    
        if self.attachments:
            att_info = self.validate_attachments_info(self.attachments)

            for att in att_info:
                email.Attachments.Add(att['path'], 
                                      att.get('display_name', None),
                                      att.get('type_attachment', OUTLOOKTYPEATTACHMENTS.RegularAttachment.value),
                                      att.get('position', 0))
        
        
        #########################################################################################
        #####               Creo el email y lo envío con o sin adjuntos                #########
        #########################################################################################
       
        email_sent = {
            'Subject': email.subject,
            'To': email.To,
            'CC': email.CC,
            'BCC': email.BCC,
            'Body': email.body if not self.is_html else email.HTMLBody,
            'Importance': self.importance_email.name,
            'IsHTML': self.is_html,
            'ReadReceiptRequested': self.read_receipt,
            'DeliveryReceiptRequested': self.delivery_receipt,
            'Attachments': [getattr(att, 'FileName', None) for att in getattr(email, 'Attachments', []) if getattr(att, 'FileName', None) is not None],
            'SentOnBehalfOfName': self.send_on_behalf if self.send_on_behalf else self.get_sender_str(email)
        }
        email.Send()
        email_sent['SentTime'] = format_datetime(datetime.now())
        df_email_sent = pd.DataFrame([email_sent])
            
        return df_email_sent
    
    
    #################################################################################################################################
    #####                          Función para enviar correos con o sin adjuntos                                           #########
    #################################################################################################################################
    
    def validate_attachments_info (self, attachments: list[EmailAttachmentInfo]):
        if not attachments:
            return None
        
        att_files = []
        for attachment in attachments:
            path = Path(attachment.file_path)
            if not path.exists() or not path.is_file():
                raise FileNotFoundError(f"El archivo adjunto {attachment.file_path} no existe o no es un archivo válido.")

            if attachment.display_name and not isinstance(attachment.display_name, str):
                raise ValueError("El nombre para mostrar del adjunto debe ser una cadena de texto.")
            
            att_file = {
                'path': str(path),
                'display_name': attachment.display_name if attachment.display_name else path.name,
                'type_attachment': attachment.type.value if attachment.type else OUTLOOKTYPEATTACHMENTS.RegularAttachment.value,
                'position': attachment.position if attachment.position and isinstance(attachment.position, int) and attachment.position >= 0 else 0
            }
            att_files.append(att_file)
            
        return att_files
    
    
    