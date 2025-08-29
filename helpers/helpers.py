import os
from dotenv import load_dotenv
from ..common import CredentialsInfoPath


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



