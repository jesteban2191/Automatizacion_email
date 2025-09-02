from .email_interface import EmailInterface
from ..common import ConnectionInfo, DataGetEmails, DataFiltersEmails, QUERYDASL, IMPORTANCEEMAIL, LOGICOPERATOR, SUBJECTPREFIX
import win32com.client
import re
import datetime
import pandas as pd
from time import time
from ..helpers import segundos_a_horas_minutos_segundos, remove_emojis
from ..common import OUTLOOKTYPERECIPENTS
import os

class OutlookEmail(EmailInterface):

    def __init__(self, conn: win32com.client.CDispatch):
        self._conn = conn # Este es el objeto Namespace autenticado
        self.tiempo_transformacion_datos_acumulado = 0
        self.tiempo_descarga_acumulado = 0
        
    
    def get_emails(self, datagetemails: DataGetEmails):
        
        
        start=time()
        #########################################################################################
        #####          Obtengo todos los parámetros para descagar los emails            #########
        #########################################################################################
        self.standard_folders = datagetemails.standard_folder
        self.custom_folder = datagetemails.custom_folder
        if self.custom_folder:
            self.custom_folder = self.custom_folder.split('/')
        self.max_emails = datagetemails.max_emails
        self.mark_as_read = datagetemails.mark_as_read
        self.page_next = datagetemails.page_next
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
        #####               Obtengo los emails según los parámetros                     #########
        #########################################################################################
        folder = self._conn.GetDefaultFolder(self.standard_folders.value)
        if self.custom_folder:
            for subfolder in self.custom_folder:
                folder = folder.Folders[subfolder]
        
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
        


        emails = []
        
        start=time()
        
        for i, message in enumerate(filtered_messages, start=self.start):
            message_class = getattr(message, 'MessageClass', None)
            is_mail = message_class == 'IPM.Note'
            is_appointment = message_class in ['IPM.Appointment', 'IPM.Schedule.Meeting.Request']

            if is_mail:
                email_data = {
                    'Tipo': 'Correo',
                    'Subject': remove_emojis(getattr(message, 'Subject', None)),
                    'Sender': f"{self.get_sender_str(message)}",
                    'To': f"{self.get_recipients_str(message, OUTLOOKTYPERECIPENTS.TO.value)}",
                    'CC': f"{self.get_recipients_str(message, OUTLOOKTYPERECIPENTS.CC.value)}",
                    'BCC': f"{self.get_recipients_str(message, OUTLOOKTYPERECIPENTS.BCC.value)}",
                    'ReceivedTime': getattr(message, 'ReceivedTime', None),
                    'SentOn': getattr(message, 'SentOn', None),
                    'Body': remove_emojis(getattr(message, 'Body', None)),
                    'HTMLBody': remove_emojis(getattr(message, 'HTMLBody', None)),
                    'Num_of_Attachments': getattr(getattr(message, 'Attachments', None), 'Count', 0),
                    'IsRead': getattr(message, 'UnRead', None) == 0,
                    'Importance': "LOW" if getattr(message, 'Importance', None) == 0 else "NORMAL" if getattr(message, 'Importance', None) == 1 else "HIGH" if getattr(message, 'Importance', None) == 2 else "Unknown",
                    'ConversationTopic': getattr(message, 'ConversationTopic', None),
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
                    'Organizer': f"{self.get_organizer_smtp(message)}",
                    'Start': getattr(message, 'Start', None),
                    'End': getattr(message, 'End', None),
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
            emails.append(email_data)
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
                query_parts.append(f"(LOWER({QUERYDASL.SUBJECT.value}) LIKE '{self.subject}')")
            else:
                query_parts.append(f"(LOWER({QUERYDASL.SUBJECT.value}) = '{self.subject}')")

        if self.sender_email:
            senders = [f"LOWER({QUERYDASL.SENDER_EMAIL.value}) = '{email.lower()}'" for email in self.sender_email]
            query_parts.append("(" + f" {self.logic_operator_between_senders.value} ".join(senders) + ")")

        if self.recipient_email:
            recipients = [f"LOWER({QUERYDASL.RECIPIENT_EMAIL.value}) = '{email.lower()}'" for email in self.recipient_email]
            query_parts.append("(" + f" {self.logic_operator_between_recipients.value} ".join(recipients) + ")")
            
        if self.sender:
            senders = [f"LOWER({QUERYDASL.SENDER_NAME.value}) LIKE '%{name.lower()}%'" for name in self.sender]
            query_parts.append("(" + f" {self.logic_operator_between_senders.value} ".join(senders) + ")")
        
        if self.recipient:
            recipients = [f"LOWER({QUERYDASL.RECIPIENT_NAME.value}) LIKE '%{name.lower()}%'" for name in self.recipient]
            query_parts.append("(" + f" {self.logic_operator_between_recipients.value} ".join(recipients) + ")")
            
        if self.cc_email:
            cc_recipients = [f"LOWER({QUERYDASL.CC_EMAIL.value}) = '{email.lower()}'" for email in self.cc_email]
            query_parts.append("(" + f" {self.logic_operator_between_recipients.value} ".join(cc_recipients) + ")")

        if self.bcc_email:
            bcc_recipients = [f"LOWER({QUERYDASL.BCC_EMAIL.value}) = '{email.lower()}'" for email in self.bcc_email]
            query_parts.append("(" + f" {self.logic_operator_between_recipients.value} ".join(bcc_recipients) + ")")
            
        if self.cc:
            cc_names = [f"LOWER({QUERYDASL.CC_NAME.value}) LIKE '%{name.lower()}%'" for name in self.cc]
            query_parts.append("(" + f" {self.logic_operator_between_recipients.value} ".join(cc_names) + ")")
        
        if self.bcc:
            bcc_names = [f"LOWER({QUERYDASL.BCC_NAME.value}) LIKE '%{name.lower()}%'" for name in self.bcc]
            query_parts.append("(" + f" {self.logic_operator_between_recipients.value} ".join(bcc_names) + ")")
            
        if self.body:
            if "%" in self.body:
                body_value = self.body.lower()
                query_parts.append(f"(LOWER({QUERYDASL.BODY_TEXT.value}) LIKE '{body_value}' OR LOWER({QUERYDASL.BODY_HTML.value}) LIKE '{body_value}')")
            else:
                body_value = f"%{self.body.lower()}%"
                query_parts.append(f"(LOWER({QUERYDASL.BODY_TEXT.value}) = '{body_value}' OR LOWER({QUERYDASL.BODY_HTML.value}) = '{body_value}')")
                
        if self.has_attachments is not None:
            has_attachments_value = 1 if self.has_attachments else 0
            query_parts.append(f"({QUERYDASL.HAS_ATTACHMENTS.value} = {has_attachments_value})")
            
        if self.is_read is not None:
            is_read_value = 1 if self.is_read else 0
            query_parts.append(f"({QUERYDASL.IS_READ.value} = {is_read_value})")
            
        if self.received_after and self.received_before:
            if self.received_after > self.received_before:
                raise ValueError("received_after no puede ser mayor que received_before")
            date_filter = (
                f"({QUERYDASL.RECEIVED_TIME.value} >= '{self.received_after.strftime('%m/%d/%Y %H:%M:%S')}' "
                f"AND {QUERYDASL.RECEIVED_TIME.value} <= '{self.received_before.strftime('%m/%d/%Y %H:%M:%S')}')"
            )
        elif self.received_after:
            date_filter = (
                f"({QUERYDASL.RECEIVED_TIME.value} >= '{self.received_after.strftime('%m/%d/%Y %H:%M:%S')}')"
            )
        elif self.received_before:
            date_filter = (
                f"({QUERYDASL.RECEIVED_TIME.value} <= '{self.received_before.strftime('%m/%d/%Y %H:%M:%S')}')"
            )
        else:
            date_filter = None

        if date_filter:
            query_parts.append(date_filter)
            
        if self.conversation_topic:
            if "%" in self.conversation_topic:
                query_parts.append(f"(LOWER({QUERYDASL.CONVERSATION_TOPIC.value}) LIKE '{self.conversation_topic.lower()}')")
            else:
                query_parts.append(f"(LOWER({QUERYDASL.CONVERSATION_TOPIC.value}) = '{self.conversation_topic.lower()}')")
        
        if self.referenceid:
            references = [f"{QUERYDASL.REFERENCE_ID.value} = '{ref}'" for ref in self.referenceid]
            query_parts.append("(" + " OR ".join(references) + ")")
            
        if self.msg_id:
            msgids = [f"{QUERYDASL.MESSAGE_ID.value} = '{self.msg_id}'"]
            query_parts.append("(" + " OR ".join(msgids) + ")")
            
        if self.importance_email:
            query_parts.append(f"({QUERYDASL.IMPORTANCE.value} = {self.importance_email.value})")
            
        if self.subject_prefix:
            query_parts.append(f"(LOWER({QUERYDASL.SUBJECT_PREFIX.value}) = '{self.subject_prefix.value.lower()}')")
            
        if not query_parts:
            query = None
        else:
            query = query_header + f" {self.logic_operator.value} ".join(query_parts) + ")"
            
        return query
        
         
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
        
    def get_recipient_smtp(self, recipient) -> str:
        address_entry = getattr(recipient, 'AddressEntry', None)
        if address_entry is not None:
            if getattr(address_entry, 'Type', None) == 'EX':
                exchange_user = getattr(address_entry, 'GetExchangeUser', lambda: None)()
                if exchange_user is not None:
                    return getattr(exchange_user, 'PrimarySmtpAddress', None)
            return getattr(recipient, 'Address', None)
        return None
    
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