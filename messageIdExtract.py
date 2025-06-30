import argparse,Microsoft365
from Microsoft365._system import System
from os import mkdir,path
if __name__=="__main__":
    _A=argparse.ArgumentParser()
    _A.add_argument("-u",required=True,metavar="name@example.com",help="Applicable Username(s)")
    _A.add_argument("-o",required=True,metavar="path/to/dir",help="Output Directory")
    _A.add_argument("--tenant-id",required=True,metavar="UUID",help="Tenant ID")
    _A.add_argument("--application-id",required=True,metavar="UUID",help="Application ID")
    _A.add_argument("--application-secret",required=True,metavar="PASS",help="Application Secret")
    _A.add_argument("--message-id",required=False,metavar="<MESS-ID>",help="Single Message ID")
    _A.add_argument("--id-file",required=False,metavar="path/to/ids.txt",help="File containing Message IDs")
    _A=_A.parse_args()

    message_ids=[]
    if _A.message_id != None:
        message_ids.append(_A.message_id)
    elif _A.id_file != None:
        with open(_A.id_file) as fileObj:
            for line in fileObj:
                message_ids.append(line.strip())
    else:
        System.output("E","A message id or a file containing message ids is required")
        raise Exception("Required Message IDs Not Present")
    if not path.exists(_A.o):
        mkdir(_A.o)
    else:
        System.output("W","Directory "+_A.o+" already exists.")
    Microsoft365.MessageIds(_A.tenant_id,_A.application_id,_A.application_secret,_A.u,_A.o,message_ids,export_content=True)