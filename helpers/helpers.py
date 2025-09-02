import os
from dotenv import load_dotenv
from ..common import CredentialsInfoPath
import re


def crear_credenciales_entorno(cred_info: CredentialsInfoPath):
    load_dotenv()
    if cred_info.credentials_path and os.path.exists(cred_info.credentials_path):
        with open(cred_info.credentials_path, "r") as f:
            cred_json_str = f.read()
        os.environ["CREDENTIALS_GMAIL_JSON"] = cred_json_str

    if cred_info.token_path and os.path.exists(cred_info.token_path):
        with open(cred_info.token_path, "r") as f:
            token_json_str = f.read()
        os.environ["TOKEN_GMAIL_JSON"] = token_json_str



def segundos_a_horas_minutos_segundos(segundos: float) -> str:
    """
    Método que se encarga de convertir un nvalor de entrada en segundos a formato hh:mm:ss
    
    Args:
        segundos (float): Valor flotante que indican una cantidad de segundos a ser convertido al formato hh:mm:ss
        
    Return:
        str: Tiempo en formato hh:mm:ss
    
    Ejemplo:
        # Ejemplo 1: 3661 segundos (1 hora, 1 minuto, 1 segundo)
        resultado = segundos_a_horas_minutos_segundos(3661)
        print(resultado)  # Salida: 01:01:01

        # Ejemplo 2: 59 segundos
        resultado = segundos_a_horas_minutos_segundos(59)
        print(resultado)  # Salida: 00:00:59

        # Ejemplo 3: 3600 segundos (exactamente 1 hora)
        resultado = segundos_a_horas_minutos_segundos(3600)
        print(resultado)  # Salida: 01:00:00
    
        """
    horas = int(segundos)//3600
    sobrante_1 = int(segundos)%3600
    minutos = sobrante_1//60
    segundos = sobrante_1%60

    if horas < 10:
        horas_str = '0' + str(horas)
    else:
        horas_str = str(horas)
    if minutos < 10:
        minutos_str = '0' + str(minutos)
    else:
        minutos_str = str(minutos)
    if segundos < 10:
        segundos_str = '0' + str(segundos)
    else:
        segundos_str = str(segundos)
    
    tiempo_str = horas_str+':'+minutos_str+':'+segundos_str
    return tiempo_str





def remove_emojis(text):
    if not isinstance(text, str):
        return text
    # Elimina caracteres fuera del plano multilingüe básico (BMP)
    return re.sub(r'[^\u0000-\uFFFF]', '', text)