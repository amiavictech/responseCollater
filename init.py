from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

from mailReader import *
from ConfigParser import SafeConfigParser


def main():
    parser = SafeConfigParser()
    parser.read('config.ini')
    store = file.Storage('token.json')
    creds = store.get()
    service = build('gmail', 'v1', http=creds.authorize(Http()))
    DATE_TIME=parser.get('RESPONSE_COLLATE','DATE_TIME')
    SUBJECT=parser.get('RESPONSE_COLLATE','SUBJECT')+DATE_TIME
    email_messages = get_email_messages(service,SUBJECT)
    store_response(email_messages)
    print("Sucessfully updated following details to excel \n ",email_messages)


if __name__ == '__main__':
  main()