from .email_interface import EmailInterface
from ..common import ConnectionInfo, DataGetEmails, DataFiltersEmails, QUERYDASL, IMPORTANCEEMAIL, LOGICOPERATOR, SUBJECTPREFIX
import win32com.client
import re
import datetime
import pandas as pd

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
        self.page_next = datagetemails.page_next
        
        
        #########################################################################################
        #####               Creo el query para obtener los emails                       #########
        #########################################################################################
        
        self.datafilter = datagetemails.filters
        if self.datafilter:
            self.query = self.create_query(self.datafilter)
        else:
            self.query = None
            
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
            
        self.total_emails = filtered_messages.Count
        self.page_number = 1 if self.page_next is None else self.page_next
        self.page_size = self.max_emails
        self.start = (self.page_number - 1) * self.page_size
        self.end = self.start + self.page_size
        filtered_messages = list(filtered_messages)  # Convierte a lista para poder hacer slicing
        filtered_messages = filtered_messages[self.start:self.end]
        
        self.page_next_result = self.page_number + 1 if self.end < self.total_emails else None


        emails = []
        
        for i, message in enumerate(filtered_messages):
            if not self.query and i >= self.max_emails:
                break
            
            email_data = {
                'Subject': message.Subject,
                'Sender': f"{message.SenderName} ({message.SenderEmailAddress})",
                'To':"; ".join([f"{recipient.Name} ({recipient.Address})" for recipient in message.Recipients]),
                'CC':"; ".join([f"{cc.Name} ({cc.Address})" for cc in message.CC]),
                'BCC':"; ".join([f"{bcc.Name} ({bcc.Address})" for bcc in message.BCC]),
                'ReceivedTime': message.ReceivedTime,
                'SentOn': message.SentOn,
                'Body': message.Body,
                'HTMLBody': message.HTMLBody,
                'Num_of_Attachments': message.Attachments.Count,
                'IsRead': message.UnRead == 0,
                'Importance': "LOW" if message.Importance == 0 else "NORMAL" if message.Importance == 1 else "HIGH" if message.Importance == 2 else "Unknown",
                'ConversationTopic': message.ConversationTopic,
                'ConversationID': message.ConversationID,
                'MessageID': message.EntryID,
                'MessageClass': message.MessageClass,
                'Categories': message.Categories,
                'Size': message.Size,
                'Attachments': [(att.FileName, att.EntryID) for att in message.Attachments],
            }
            emails.append(email_data)
            
        df_emails = pd.DataFrame(emails)
        
        result = {
            "data": df_emails,
            "page_number": self.page_number,
            "page_size": self.page_size,
            "total_emails": self.total_emails,
            "has_next_page": self.end < self.total_emails,
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
        
         
