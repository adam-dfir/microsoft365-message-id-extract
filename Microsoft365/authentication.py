import requests,json,time
from ._system import System
class Authentication:
    def __init__(self,tenant_id,application_id,application_secret):
        if hasattr(self,"expiration") and hasattr(self,"headers"):
            if self.expiration>int(time.time()):
                return True
        _t=requests.post(f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",data={"client_id":application_id,"scope":"https://graph.microsoft.com/.default","client_secret":application_secret,"grant_type":"client_credentials"})
        if _t.status_code==200:
            _t1=json.loads(_t.text)
            self.headers={"Authorization":f"Bearer {_t1['access_token']}"}
            self.expiration=int(time.time())+_t1['expires_in']
            System.output("S","Authenticated to Microsoft365")
        else:
            System.output("W","Error authenticating to Microsoft365")
            raise Exception("Error while attempting to retrieve the Microsoft365 tokens. Invalid application ID/secret?")