from email.message import EmailMessage
import base64
import sys  
import json
import logging
import msal
import requests


config = {
    "authority": "https://login.microsoftonline.com/eb8db77e-65e0-4fc3-.........",
    "client_id": "18bb3888-dad0-4997-96b1-........",
    "scope": ["https://graph.microsoft.com/.default"],
    "secret": ".........",
    "tenant-id": "eb8db77e-65e0-4fc3-b967-b.........."
}

app = msal.ConfidentialClientApplication(config['client_id'], authority=config['authority'],
                                             client_credential=config['secret'])
result = app.acquire_token_silent(config["scope"], account=None)

if not result:
    logging.info("No suitable token exists in cache. Let's get a new one from AAD.")
    result = app.acquire_token_for_client(scopes=config["scope"])

sender_email = "sender@domain.onmicrosoft.com"
receiver_email = "receiver_email@domain.onmicrosoft.com"


def create_message(sender, to, subject, message_text):

    message = EmailMessage()
    message.set_content(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes())
    return raw.decode()

messageToSend = create_message(sender_email,receiver_email,"test subject","test 123")

print(messageToSend)
 
endpoint = f'https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail'
r = requests.post(endpoint, data=messageToSend,
                      headers={'Authorization': 'Bearer ' + result['access_token'], "Content-Type": "text/plain"})
print(r.status_code)
print(r.text)

