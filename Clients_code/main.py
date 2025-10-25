from Automatizacion_email import Orchestrator_email, ConnectionInfo, DataGetEmails, DataFiltersEmails, OutlookStandarFolders, IMPORTANCEEMAIL, LOGICOPERATOR, SUBJECTPREFIX, EmailContext, DataDownloadAttachments
import pandas as pd
from datetime import datetime, timezone
import tzlocal

def intialize():
    conn_info = ConnectionInfo(email_provider='OUTLOOK')

    email_service = Orchestrator_email(conn_info)
    return email_service
##received_after= '01/01/2025'
##subject='%Insumos Base Certificada gobRed%', 
def intialize_data():
    local_tz = tzlocal.get_localzone()
    # fec_ini = datetime(2025, 10, 1, 0, 0, 0, tzinfo=timezone.utc).astimezone(local_tz).replace(tzinfo=None)
    # fec_fin = datetime(2025, 10, 16, 0, 0, 0, tzinfo=timezone.utc).astimezone(local_tz).replace(tzinfo=None)
    fec_ini = "01/10/2025"
    fec_fin = "16/10/2025"

    filters = DataFiltersEmails(subject='%Insumos Base Certificada gobRed%', received_after= fec_ini, received_before= fec_fin)
    settings_download = DataDownloadAttachments(download_folder="C:/Users/jueriver/Documents/Base_zonas_gob_red_prueba", overwrite= True, only_extensions=['.xlsx'], name_subfolder_per_email="Base_zonas_{reciveddate}")
    dataemails = DataGetEmails(store_folder= "jueriver@bancolombia.com.co", standard_folder=OutlookStandarFolders.INBOX, filters=filters, download_attachments= False, attachments_settings= settings_download, mark_as_read=True, page_size= 500)
    return dataemails
def main():
    email_service = intialize()
    dataemails = intialize_data()

    result = email_service.get_emails(dataemails, return_all_pages=True)

    df_emails = result['data']
    print(df_emails.head())
    page_number = result['page_number']
    page_size = result['page_size']
    total_emails = result['total_emails']
    has_more = result['has_more']
    page_next = result['page_next']
    
    df_emails.to_excel('emails_outlook_zonas.xlsx', index=False)
    
    print(f"Página: {page_number}, Tamaño de página: {page_size}, Total de emails: {total_emails}, Hay más páginas: {has_more}, Página siguiente: {page_next}")
    
    
    
if __name__ == "__main__":
    main()