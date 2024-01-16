import smtplib
from email.message import EmailMessage
from ssl import create_default_context
from string import Template
from pathlib import Path
from csv import DictReader


def send_email(email_senders_file, email_receivers_file, attachments_file):
    try:
        with open(email_senders_file, "r") as file:
            reader1 = DictReader(file)
            users = list(reader1)
        for user in users:
            email_sender = user.get('email_sender')
            email_acc_password = user.get('password')
            subject = user.get('email_subject')

        with open(email_receivers_file, 'r') as file1:
            reader2 = DictReader(file1)
            users1 = list(reader2)
        for user1 in users1:
            name = user1.get('name')
            receivers_email = user1.get('email_receiver')

            html = Template(Path('F:\PRANAV\Email\sample.html').read_text())
            email = EmailMessage()
            email['from'] = email_sender
            email['to'] = receivers_email
            email['subject'] = subject
            email.set_content(html.substitute({'name': name}), "html")
            context = create_default_context()

            with open(attachments_file, 'r') as file2:
                reader3 = DictReader(file2)
                attachmentsfile = list(reader3)
            for filename in attachmentsfile:
                attachment_path = filename.get('file-path')
                with open(attachment_path, 'rb') as f:
                    attachment_data = f.read()
                filename = Path(attachment_path).name
                email.add_attachment(
                    attachment_data, maintype='application', subtype='octet-stream', filename=filename)

            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
                smtp.ehlo()
                smtp.login(email_sender, email_acc_password)
                smtp.send_message(email)

    except Exception as err:
        print('Hmm Something Went wrong!')
        print(err)


send_email('F:\PRANAV\Email\email_senders.csv',
           'F:\PRANAV\Email\email_receivers.csv', 'F:\PRANAV\Email\Attachments.csv')
