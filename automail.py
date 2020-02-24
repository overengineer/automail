#!/usr/bin/python3
import openpyxl, os.path, yaml
from glob import glob
from getpass import getpass
from mailer import Mailer, Message
from munch import Munch

# Load configuration
with open('config.yaml') as f:
    config_dict = yaml.safe_load(f)
config_dict = {k:v.strip() for k,v in config_dict.items()}
config = Munch(config_dict)


# Load receivers
wb = openpyxl.load_workbook(
        filename=config.receivers_path, 
        data_only=True, 
        read_only=True
    )
rows = [row[:2] for row in list(wb.active.values)[1:]]
receivers = [(name.strip(), mail.strip()) for name, mail in rows if name and mail]

# Load template
with open(config.template_path) as f:
    template = f.read()

# Inline styles 
import premailer
template = premailer.transform(template)

# Clean HTML
import lxml.html
from lxml.html.clean import Cleaner
cleaner = Cleaner()
cleaner.kill_tags = ['style', 'script']
page = cleaner.clean_html(lxml.html.fromstring(template))
assert not page.xpath('//style'), 'style'
assert not page.xpath('//script'), 'script'
template = lxml.html.tostring(page).decode('utf-8')

# Send mails
sender = Mailer('smtp.yandex.com.tr', port='465', use_ssl=True)
sender.login(config.user_mail, getpass('Password: '))
print('start')
for receiver_name, receiver_mail in receivers:
    try:
        message = Message(From=config.user_mail,
                  To=receiver_mail,
                  charset="utf-8")
        attachment_path = glob(os.path.join(config.attachments_path,receiver_name+'.*'))[0]
        message.Subject = config.subject   
        message.Html = template.format(NAME=receiver_name)
        message.attach(attachment_path)
        sender.send(message)
    except Exception as ex:
        print(receiver_name, receiver_mail, 'Failed\n', ex)
print('Done')

