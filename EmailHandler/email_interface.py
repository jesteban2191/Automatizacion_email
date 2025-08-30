from typing import Protocol, Optional
from ..common import DataGetEmails
import pandas as pd

class EmailInterface(Protocol):
    conn: object  # Puede ser namespace (Outlook) o service (Gmail)
    datagetemails: object
    datafilter: Optional[object]
    
    def get_emails(self, datagetemails: object) -> pd.DataFrame:
        pass
    
    def create_query(self, datafilters: object) -> str:
        pass
    
    