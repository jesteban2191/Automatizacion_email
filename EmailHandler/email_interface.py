from typing import Protocol
from ..common import DataGetEmails
import pandas as pd

class EmailInterface(Protocol):
    conn: object  # Puede ser namespace (Outlook) o service (Gmail)
    datagetemails: DataGetEmails
    
    def get_emails(self, datagetemails: DataGetEmails) -> pd.DataFrame:
        pass
    
    
    