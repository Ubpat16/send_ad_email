import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas

# Using Pandas to open the list and obtain data rows
entrance_list = pandas.read_csv('admission_list.csv')
for index in range(len(entrance_list)):
    password = ''
    name = entrance_list['Name'][index]
    email_address = entrance_list['email_address'][index]
    child = entrance_list['first_name'][index]
    file = f'f_letters\\{name}.pdf'

    # An instance of MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = 'topfaithlegacycollege@topfaith.sch.ng'
    msg['To'] = email_address
    msg['Subject'] = '2022/2023 PROVISIONAL ADMISSION FOR ' + name
    body = f'Dear Parent,\n\nCongratulations on your choice of Topfaith Legacy College Ibiakpan for your {child}. Forwarded herewith is your ward Admission/Result notification letter for the entrance examination held on Saturday, April 2, 2022.\n\n Thank you.\n\n TOPFAITH SCHOOLS'

    # Attaching body with msg
    msg.attach(MIMEText(body, 'plain'))

    # files to be sent to their respective emails
    filename = f'{name}.pdf'
    attachment = open(file, 'rb')

    # instance of MIMEBase
    mime_base = MIMEBase('application', 'octet-stream')

    # Setting payload into encoded form
    mime_base.set_payload(attachment.read())

    # Encode into base64
    encoders.encode_base64(mime_base)

    mime_base.add_header('Content-Disposition', 'attachment; filename= %s' % filename)
    msg.attach(mime_base)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', port=465, context=context) as conn:
        print(f'Mail sending to {name}')
        conn.login(user=msg['From'], password=password)
        text = msg.as_string()
        conn.sendmail(from_addr=msg['From'], to_addrs=email_address, msg=text)
        print(f'Mail sent to {name}')


