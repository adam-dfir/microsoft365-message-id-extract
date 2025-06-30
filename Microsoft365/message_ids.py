import requests,json,uuid,base64,re,csv
from .authentication import Authentication
from ._system import System
from os import mkdir
class MessageIds(Authentication):
    def __init__(self,tenant_id,application_id,application_secret,user_id,export_dir,message_ids,export_content=True):
        super().__init__(tenant_id,application_id,application_secret)
        messages=[]
        for messageId in message_ids:
            message=self.get_message(messageId,user_id)
            _id=str(uuid.uuid4())
            abs_path=export_dir+"/"+re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", message['subject'])+" - "+_id
            mkdir(abs_path)
            if export_content == True:
                self.export_message(message,abs_path)
                self.export_attachments(message['id'],user_id,abs_path)
                System.output("S","Exported Message Content and Attachments for "+message['subject'])
            messages.append(message)
        self.export_metadata(messages,export_dir)
    def get_message(self,messageId,user_id):
        _t=requests.get(f"https://graph.microsoft.com/v1.0/users/{user_id}/messages?$filter=internetMessageId eq '{messageId}'",headers=self.headers)
        if _t.status_code==200:
            message=json.loads(_t.text)['value'][0]
            _r={}
            for field in message:
                if field == "sender" or field == "from":
                    _r[field]=message[field]['emailAddress']['address']
                elif field == "toRecipients" or field == "ccRecipients" or field == "bccRecipients" or field == "replyTo":
                    _t1=[]
                    for _ in message[field]:
                        _t1.append(_['emailAddress']['address'])
                    _r[field]=";".join(_t1)
                else:
                    _r[field]=message[field]
            System.output("I","Found Email "+message['subject'])
            return _r
    def export_attachments(self,id,user_id,output_directory):
        _t=requests.get(f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{id}/attachments",headers=self.headers)
        if _t.status_code==200:
            _r=[]
            for attachment in json.loads(_t.text)['value']:
                _r.append(attachment['name'])
                with open(output_directory+"/Attachment - "+attachment['name'], "wb") as objFile:
                    _d=base64.b64decode(attachment['contentBytes'])
                    objFile.write(_d)
            return _r
    def export_message(self,message,output_directory):
        with open(output_directory+"/Email Content.html","w",encoding="utf-8") as objFile:
            objFile.write(message['body']['content'])
    def export_metadata(self,messages,output_directory):
        _data=[{k: v for k, v in row.items() if k != "body" and k != "bodyPreview"} for row in messages]
        _keys=sorted({k for row in _data for k in row})
        with open(output_directory+"/Email Metadata.csv","w",newline="",encoding="utf-8-sig") as csvFile:
            _w=csv.DictWriter(csvFile,fieldnames=_keys,delimiter=",")
            _w.writeheader()
            _w.writerows(_data)
        System.output("S","Exported Message Metadata")