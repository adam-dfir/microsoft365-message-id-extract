# Microsoft365 Message ID Email Extraction
> Simple script to extract emails, attachments, and metadata using the Graph API to bypass date limitations implemented through message tracing and other email extraction methods.
## API Setup
You need access to the Microsoft Graph API. To register an application:
1. Register an aplication within Entra and create an application secret
2. Grant admin content to the `Mail.Read` permission
## Message IDs
This script searches for message ids to extract email content and attachments.  Message Ids can be found in logs (e.g. MailItemsAccessed) or in email headers.  An example message id is `<ABC123123123123123123123@ABC123123123.namprd1.prod.outlook.com>`
### Usage
```
python messageIdExtract.py -h
```
```
python .\messageIdExtract.py --tenant-id abc123ab-abc1-abc1-abc1-abc123123123 --application-id xyz12312-xyz12-xyz12-xyz12-xyz123123123 --application-secret ABCD_abc123abc123abc123abc123 --id-file "message_ids.txt" -u email@example.com -o export_directory
```
### Requirements
```
pip install requests
```