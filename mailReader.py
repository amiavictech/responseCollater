from openpyxl import load_workbook
from apiclient import errors
import base64
import email
import re

def get_email_messages(service, subjectFilter):
    try:
        response = service.users().messages().list(userId='me', labelIds=['INBOX'], q=subjectFilter).execute()
        messages = []

        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId='me', labelIds=['INBOX'], q=subjectFilter, pageToken=page_token).execute()
            messages.extend(response['messages'])
        fullMessages = []

        for msg in messages:
            mymsg = {}
            details = service.users().messages().get(userId='me', id=msg['id'],format='full').execute()
            for headerProps in details['payload']['headers']:
                if headerProps['name']=='From':
                    #mymsg['email']=re.search('(?<=\<).*?(?=\>)',headerProps['value']).group()
                    mymsg['email']=headerProps['value']
                    print(headerProps['value'])
                elif headerProps['name']=='Subject':
                    splitStr = headerProps['value'].split('#')
                    mymsg['status'] = splitStr[1]
                    mymsg['adult_count'] = splitStr[2]
                    mymsg['kids_count'] = splitStr[3]
            fullMessages.append(mymsg);
            #
            # mymsg['value']=GetMimeMessage(service,'me',msg['id'])

        return fullMessages
    except errors.HttpError, error:
        print 'An error occurred: %s' % error

def GetMimeMessage(service, user_id, msg_id):
  """Get a Message and use it to create a MIME Message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: The ID of the Message required.

  Returns:
    A MIME Message, consisting of data from Message.
  """
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id,format='full').execute()

    mymsg = message['payload']['parts'][0]['body']['data']
    #print ()

    msg_str = base64.urlsafe_b64decode(mymsg.encode('ASCII'))

    #mime_msg = email.message_from_string(msg_str)
    return msg_str
  except errors.HttpError, error:
    print 'An error occurred: %s' % error

def store_response(responseMails):
    wb = load_workbook('ResponseStatus.xlsx')
    sheet = wb['Sheet1']

    email_column = sheet['E']
    response_status = sheet['I']
    adl_cnt = sheet['J']
    kid_cnt = sheet['K']

    # Print the contents
    for response in responseMails:
        for x in xrange(len(email_column)):
            if x > 0 and email_column[x].value == response['email']:
                response_status[x].value=response['status']
                adl_cnt[x].value=response['adult_count']
                kid_cnt[x].value=response['kids_count']
                break
    wb.save('ResponseStatus.xlsx')
    print('Saved to excel')