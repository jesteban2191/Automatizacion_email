from typing import Protocol
from ..common import ConnectionInfo

class AuthInterface(Protocol):
    conn_info: ConnectionInfo
    
    def authenticate(self):
        pass