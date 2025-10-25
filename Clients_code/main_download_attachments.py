from Automatizacion_email import AuthContext, ConnectionInfo, CredentialsInfoPath, PermisosGmail, DataGetEmails, DataFiltersEmails, OutlookStandarFolders, IMPORTANCEEMAIL, LOGICOPERATOR, SUBJECTPREFIX, EmailContext, DataDownloadAttachments
import pandas as pd

def intialize():
    conn_info = ConnectionInfo(email_provider='OUTLOOK')

    auth_context = AuthContext(conn_info)
    auth_context.authenticate()
    namespace = auth_context.get_namespace()
    return namespace

def intialize_data():
    filters = DataFiltersEmails( subject='%Base zonas%', received_after= '01/01/2025')
    datadownloadattachments = DataDownloadAttachments(download_folder= "C:/Users/jueriver/Documents/Base_zonas_gob_red", mark_as_read=True, only_extensions=['.xlsx'],filters=filters, name_subfolder_per_email="Base_zonas_{reciveddate}")
    return datadownloadattachments

def main():
    namespace = intialize()
    print(f"namespace type: {type(namespace)}")
    email_service = EmailContext(namespace)
    datadownloadattachments = intialize_data()
    df_result = email_service.download_attachments(datadownloadattachments)
    df_result.to_excel('attachments_outlook_zonas.xlsx', index=False)

if __name__ == "__main__":
    main()