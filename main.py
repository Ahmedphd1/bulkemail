import configparser
from sendmail import *
import json
import time

config = configparser.ConfigParser()
config.read('config.ini')

attachments = json.loads(config['variables']['filer'])

for email in getemaillist():
    sendmessage(email, createmessage(attachments))

    print(f"Email sent to: {email}")

    time.sleep(10)
