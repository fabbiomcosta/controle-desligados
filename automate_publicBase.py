import sys, io, os, json, argparse, slack
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload
from google.cloud import storage
from datetime import date
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv  


# Define the program description
TEXT = 'This is a script to download xlsx files on google drive, convert it to json and save it on GCP Storage'

# Version
VERSION = 'Version 1.3'

# API permission type
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# Service account key
SERVICE_ACCOUNT_FILE = 'service-accounts-273522.json'

# The GCP bucket name 
BUCKET = "siem-prod"

TODAY = date.today()
# Query to search for files on the current date.
MIMETYPE = "mimeType != 'application/vnd.google-apps.folder' and (modifiedTime > '{}')".format(TODAY)

# Slack API
CHANNEL = '#teste_save_files_bucket'

def download_from_drive():

    # Authentication with the service account
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

        service = build('drive', 'v3', credentials=creds)
    except Exception as e:
        print("Check the Authentication session")
        print( "Error: %s" % str(e) )
        sys.exit(1)

    # Search list of files with the current date
    try:
        results = service.files().list(q=MIMETYPE, fields="nextPageToken, files(id, name)").execute()
        files = results.get('files', [])
        for item in files:
            if "Base" in item['name']:
                file_id = item['id']
                file_name = item['name']
                print("Name: %s" % (file_name))
                print("ID: %s" % (file_id))
    except Exception as e:
        print('No files found.')
        print( "Error: %s" % str(e) )
        sys.exit(1)

    # Download the current file
    try:
        request = service.files().get_media(fileId=file_id)
        fh = io.FileIO(file_name, mode='wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print("Download %d%%." % int(status.progress() * 100))
    except Exception as e:
        print('Check the Authentication session')
        print( "Error: %s" % str(e) )
        sys.exit(1)

    return file_name

def convert_xlxs2Json(file_name):
    try:
        # save in json only thoses columns: B,E,I,J,K,M,O,P,R,W and X
        read_file = pd.read_excel (file_name, index_col=None, na_values=['NA'], usecols = 'B,E,I,J,K,M,O,P,R,W,X')
        result = read_file.to_json(orient='records', date_format='iso   ')
        parsed = json.loads(result)
        file_name = TODAY.strftime("desligados_%d%m%Y.json")
        with open(file_name, 'w', encoding='utf-8') as jsonf:
            jsonf.write(json.dumps(parsed, ensure_ascii=False, indent=4))
        return file_name
    except Exception as e:
        print('Check the convert json session')
        print( "Error: %s" % str(e) )
        sys.exit(1) 

def upload_to_GCP(file_name):
    try:
        client = storage.Client.from_service_account_json(SERVICE_ACCOUNT_FILE)
        bucket = client.get_bucket(BUCKET)
        blob = bucket.blob(file_name)
        blob.upload_from_filename(file_name)
    except Exception as e:
        print('Check the upload json session')
        print( "Error: %s" % str(e) )
        sys.exit(1) 

def send_msg_slack (mensagem, ch):
    
    # Load SLACK_TOKEN_STONE environment variables from the .env file
    env_path = Path('.') / '.env'
    load_dotenv(dotenv_path=env_path)
    try:
        client = slack.WebClient(token=os.environ['SLACK_TOKEN_STONE'])
        client.chat_postMessage(channel=ch, text=mensagem)
        return True
    except Exception as e:
        print("Send message slack session")
        print( "Error: %s" % str(e) )
        return False

def delete_files(file_json, file_xlsx):
    os.remove(file_json)
    os.remove(file_xlsx)
    


if __name__ == '__main__':
    # initiate the parser
    parser = argparse.ArgumentParser(description = TEXT)
    parser.add_argument("-v", "--version", help="show program version", action="store_true")
    parser.add_argument("-b", "--base", help=" type of base: [desligados|leak]")
    parser.add_argument("-d", "--delete", help="delete all json an xlsx", action="store_true")

    args = parser.parse_args() 

    if args.version:
        print(VERSION)
    elif args.base:
        if args.base == "desligados":
            file_xlsx = download_from_drive()
            file_json = convert_xlxs2Json(file_xlsx)
            #upload_to_GCP(file_json)

            message = "[Automate public base - Desligados]\n *Arquivo:* " + file_json
            #send_msg_slack(message, CHANNEL)

            if args.delete:
                delete_files(file_json,file_xlsx)

        if args.base == "leak":
            file_xlsx = download_from_drive()
            file_json = convert_xlxs2Json(file_xlsx)
            upload_to_GCP(file_json)

            message = "[Automate public base - Leak]\n *Arquivo:* " + file_json
            #send_msg_slack(message, CHANNEL)    
                        
            if args.delete:
                delete_files(file_json,file_xlsx)                        

        else:
            print ("required arguments: automate_publicBase.py -h | -- help")
    else:
        print ("required arguments: automate_publicBase.py -h | -- help")