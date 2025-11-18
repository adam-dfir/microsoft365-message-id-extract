############################################################################################################
## November 2025
## Extract Emails by ID
## This script will extract email messages via ID using Microsoft's Graph API
## Requirements: requests
## An application is required for API access that has admin consent granted on the "Mail.Read" permission
## Usage: extract-emails-by-id.py --app-id UUID --app-secret SECR --tenant-id UUID --identity email@example.com --output path/to/folder [--ids-file path/to/file.txt]
## The Message IDs to search for must be provided via standard input (e.g., piped from a file).
############################################################################################################

import requests, argparse, json, time, uuid, base64, re, csv, sys
from os import mkdir, path

def colored_output(type, message):
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    _C = {
        "W": ("[!] ", WARNING),
        "S": ("[+] ", OKGREEN),
        "I": ("[*] ", OKCYAN),
        "E": ("[-] ", FAIL)
    }
    if type in _C:
        code, color = _C[type]
        print(f"{color}{code}{ENDC}{message}")
    else:
        print(f"{OKCYAN}[?]{ENDC} {message}")
class Authentication:
    def __init__(self, tenant_id, application_id, application_secret):
        self.session = requests.Session()
        if hasattr(self, "expiration") and hasattr(self, "headers"):
            if self.expiration > int(time.time()):
                return
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": application_id,
            "scope": "https://graph.microsoft.com/.default",
            "client_secret": application_secret,
            "grant_type": "client_credentials"
        }
        _t = requests.post(url, data=data)
        if _t.status_code == 200:
            _t1 = json.loads(_t.text)
            self.headers = {"Authorization": f"Bearer {_t1['access_token']}"}
            self.expiration = int(time.time()) + _t1['expires_in']
            self.session.headers.update(self.headers)
            colored_output("S", "Authenticated to Microsoft365")
        else:
            colored_output("E", f"Error authenticating to Microsoft365. Status: {_t.status_code}. Check your application ID, tenant ID, and application secret.")
            raise Exception("Authentication failed. Invalid application ID/secret or tenant ID.")

class MessageExtract(Authentication):
    BASE_URL = "https://graph.microsoft.com/v1.0/"
    def __init__(self, tenant_id, application_id, application_secret, user_id, export_dir, message_ids, export_content=True):
        super().__init__(tenant_id, application_id, application_secret)
        messages = []
        total_ids = len(message_ids)
        found_count = 0
        colored_output("I", f"Processing {total_ids} message ID(s) for user {user_id}...")
        for messageId in message_ids:
            message = self.get_message(messageId, user_id)
            if not message:
                continue
            found_count += 1
            _id = str(uuid.uuid4())
            safe_subject = re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", message['subject'])
            abs_path = path.join(export_dir, f"{safe_subject} - {_id}")
            mkdir(abs_path)
            if export_content:
                self.export_message(message, abs_path)
                self.export_attachments(message['id'], user_id, abs_path)
                colored_output("S", f"Exported Message Content and Attachments for {message['subject']}")
            messages.append(message)
        colored_output("S", "--- Search Summary ---")
        colored_output("S", f"Total IDs processed: {total_ids}")
        colored_output("S", f"Messages Found:      {found_count}")
        colored_output("W", f"Messages Not Found:  {total_ids - found_count}")
        colored_output("S", "----------------------")
        if messages:
            self.export_metadata(messages, export_dir)
        else:
            colored_output("W", "No messages were found, skipping metadata export.")

    def _process_message(self, message):
        _r = {}
        for field in message:
            if field in ["sender", "from"]:
                _r[field] = message[field]['emailAddress']['address']
            elif field in ["toRecipients", "ccRecipients", "bccRecipients", "replyTo"]:
                _t1 = [item['emailAddress']['address'] for item in message[field]]
                _r[field] = ";".join(_t1)
            else:
                _r[field] = message[field]
        colored_output("S", f"Found Email '{message['subject']}' (ID: {message.get('internetMessageId', 'N/A')})")
        return _r

    def get_message(self, messageId, user_id):
        folder_endpoints = {
            "Inbox": f"users/{user_id}/messages",
            "Sent Items": f"users/{user_id}/mailFolders/sentitems/messages",
            "Deleted Items": f"users/{user_id}/mailFolders/deleteditems/messages",
            "Junk Email": f"users/{user_id}/mailFolders/junkemail/messages",
            "Archive": f"users/{user_id}/mailFolders/archive/messages",
            "Drafts": f"users/{user_id}/mailFolders/drafts/messages",
            "Recoverable Items (Deletions)": f"users/{user_id}/mailFolders/recoverableitemsdeletions/messages"
        }
        colored_output("I", f"Searching for Message ID: {messageId}")
        for folder_name, endpoint_suffix in folder_endpoints.items():
            url = f"{self.BASE_URL}{endpoint_suffix}?$filter=internetMessageId eq '{messageId}'"
            colored_output("I", f"  -> Checking in folder: {folder_name}...")
            _t = self.session.get(url)
            if _t.status_code == 200:
                data = json.loads(_t.text)
                if data.get('value'):
                    colored_output("S", f"     Found in {folder_name}.")
                    return self._process_message(data['value'][0])
            else:
                colored_output("W", f"     API error checking {folder_name}. Status: {_t.status_code}. Skipping folder.")
        colored_output("W", f"Could not find message with ID: {messageId} in any of the checked folders.")
        return None

    def export_attachments(self, id, user_id, output_directory):
        url = f"{self.BASE_URL}users/{user_id}/messages/{id}/attachments"
        _t = self.session.get(url)
        if _t.status_code == 200:
            attachments = json.loads(_t.text).get('value', [])
            if not attachments:
                colored_output("I", "   -> No attachments found for this message.")
                return []
            colored_output("I", f"   -> Found {len(attachments)} attachment(s). Exporting...")
            for attachment in attachments:
                attachment_filename = path.join(output_directory, f"Attachment - {attachment['name']}")
                try:
                    _d = base64.b64decode(attachment['contentBytes'])
                    with open(attachment_filename, "wb") as objFile:
                        objFile.write(_d)
                except Exception as e:
                    colored_output("E", f"     Failed to export attachment '{attachment['name']}': {e}")
            return [a['name'] for a in attachments]
        else:
            colored_output("W", f"   -> Could not retrieve attachments. Status: {_t.status_code}")
            return []

    def export_message(self, message, output_directory):
        if message.get('body', {}).get('content'):
            output_file = path.join(output_directory, "Email Content.html")
            with open(output_file, "w", encoding="utf-8") as objFile:
                objFile.write(message['body']['content'])
        else:
            colored_output("W", "   -> Message body content not found or empty.")

    def export_metadata(self, messages, output_directory):
        _data = [{k: v for k, v in row.items() if k not in ["body", "bodyPreview"]} for row in messages]
        _keys = sorted({k for row in _data for k in row})
        output_file = path.join(output_directory, "Email Metadata.csv")
        with open(output_file, "w", newline="", encoding="utf-8-sig") as csvFile:
            _w = csv.DictWriter(csvFile, fieldnames=_keys, delimiter=",")
            _w.writeheader()
            _w.writerows(_data)
        colored_output("S", "Exported Message Metadata")

if __name__ == "__main__":
    arguments = argparse.ArgumentParser(description="Extract Emails by ID")
    arguments.add_argument("--app-id", required=True, help="The application (client) ID")
    arguments.add_argument("--tenant-id", required=True, help="The tenant ID")
    arguments.add_argument("--app-secret", required=True, help="The application (client) Secret")
    arguments.add_argument("--identity", required=True, help="The affected identity/email address (e.g., user@example.com)")
    arguments.add_argument("--output", required=True, help="The output directory")
    arguments.add_argument("--ids-file", help="Path to a file containing Message IDs (one per line). If not provided, IDs are read from stdin.", default=None)
    arguments = arguments.parse_args()
    if not path.isdir(arguments.output):
        mkdir(arguments.output)
        colored_output("I", f"Created output directory: {arguments.output}")
    message_ids = []
    if arguments.ids_file:
        try:
            with open(arguments.ids_file, 'r') as f:
                message_ids = [line.strip() for line in f if line.strip()]
            colored_output("I", f"Read {len(message_ids)} message ID(s) from {arguments.ids_file}.")
        except FileNotFoundError:
            colored_output("E", f"Error: Message IDs file not found at {arguments.ids_file}.")
            exit()
    else:
        colored_output("I", "Reading Message IDs from standard input (stdin)...")
        try:
            message_ids = [line.strip() for line in sys.stdin if line.strip()]
            if not message_ids:
                 colored_output("E", "No Message IDs provided via standard input.")
                 exit()
            colored_output("I", f"Read {len(message_ids)} message ID(s) from stdin.")
        except Exception as e:
            colored_output("E", f"Error reading from stdin: {e}")
            exit()
    if not message_ids:
        colored_output("W", "No valid Message IDs were provided. Exiting.")
        exit()
    try:
        MessageExtract(
            tenant_id=arguments.tenant_id,
            application_id=arguments.app_id,
            application_secret=arguments.app_secret,
            user_id=arguments.identity,
            export_dir=arguments.output,
            message_ids=message_ids
        )
        colored_output("S", "Script execution complete.")
    except Exception as e:
        colored_output("E", f"A critical error occurred: {e}")
