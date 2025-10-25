from Automatizacion_email import Orchestrator_email, ConnectionInfo, DataGetEmails, DataFiltersEmails, OutlookStandarFolders, IMPORTANCEEMAIL, LOGICOPERATOR, SUBJECTPREFIX, EmailContext, DataDownloadAttachments, DataSendEmail
import pandas as pd
from helper.helper import Helper
from datetime import datetime
from dateutil.relativedelta import relativedelta

ih = Helper(dsn="impala_prod")

def intialize():
    conn_info = ConnectionInfo(email_provider='OUTLOOK')

    email_service = Orchestrator_email(conn_info)
    return email_service

def initialize_data():
    file_path = "correo.html"
    fecha_actual = datetime.now()
    año = fecha_actual.year
    mes = fecha_actual.month
    dia = fecha_actual.day
    fecha_filtro = fecha_actual - relativedelta(months=6)
    año_ant = fecha_filtro.year
    mes_ant = fecha_filtro.month
    with open(file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()
        
    query1 = f"""select 
                    ingestion_year, 
                    ingestion_month, 
                    ingestion_day, 
                    count(*) as total_registros 
                from resultados_clientes_personas_y_pymes.ejecutivos_zonas_gobierno_de_red
                where cast(ingestion_year as bigint)*10000 + cast(ingestion_month as bigint)*100 >= {año_ant*10000 + mes_ant*100}
                group by ingestion_year, ingestion_month, ingestion_day
                order by ingestion_year desc, ingestion_month desc, ingestion_day desc;"""
    
    query2 = f"""select 
                    ingestion_year, 
                    ingestion_month, 
                    ingestion_day, 
                    count(*) as total_registros 
                from resultados_clientes_personas_y_pymes.BASE_CERTIFICADA_GOBIERNO_RED
                where cast(ingestion_year as bigint)*10000 + cast(ingestion_month as bigint)*100 >= {año_ant*10000 + mes_ant*100}
                group by ingestion_year, ingestion_month, ingestion_day
                order by ingestion_year desc, ingestion_month desc, ingestion_day desc;"""
                
    df_base_zonas = ih.obtener_dataframe(query1)
    df_base_certificada = ih.obtener_dataframe(query2)
    html_base_zonas = df_base_zonas.to_html(index=False, justify= "center")
    html_base_zonas = html_base_zonas.replace("<table ", "<table align='center' ")
    html_base_certificada = df_base_certificada.to_html(index=False, justify= "center")
    html_base_certificada = html_base_certificada.replace("<table ", "<table align='center' ")
    print(html_base_certificada)
    html_content = html_content.replace("{tabla_html_Base_zonas_gob_red}", html_base_zonas)
    html_content = html_content.replace("{tabla_html_Base_certificada_gob_red}", html_base_certificada)
    subject = f"Ingesta Exitosa {año*10000+mes*100+dia} - Base Zonas Gobierno de Red y Base Certificada Gobierno de Red"
    recipients = ["jdavid@bancolombia.com.co"]
    cc = ["LAUHERRE@bancolombia.com.co", "YVZAPATA@bancolombia.com.co", "jueriver@bancolombia.com.co"]
    datasendemail = DataSendEmail(subject=subject, body=html_content, to_recipients_email=recipients, cc_recipients_email=cc)
    
    return datasendemail
    
    
def main():
    email_service = intialize()
    
    datasendemail = initialize_data()

    result = email_service.send_email(datasendemail)
    
    print(result)
    
    result.to_excel('resultado_envio_email.xlsx', index=False)


if __name__ == "__main__":
    main()