import imaplib, email
import docx

user = 'pythonemailtest007@gmail.com'
pwd = 'your pwd'
imp_url = 'imap.gmail.com'

print('Python prgm')
con = imaplib.IMAP4_SSL(imp_url)
con.login(user,pwd)
con.select('inbox')

type, data = con.search(None, 'ALL')

doc = docx.Document()

mail_ids = data[0]
id_list = mail_ids.split()
first_mail_id = int(id_list[0])
last_mail_id = int(id_list[-1])
print(last_mail_id)
print(first_mail_id)
for i in range(last_mail_id, first_mail_id, -1):
    typ, data = con.fetch(b'2', '(RFC822)')
    for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_string(response_part[1].decode('utf-8'))
            email_subject = msg['subject']
            email_from = msg['from']
            print('From : ' + email_from + '\n')
            print('Subject : ' + email_subject + '\n')
            for part in msg.walk():
                #print('inside mes walk')
                if part.get_content_type() == 'text/plain':
                    body = part.get_payload(decode=True)
                    print(body)
                    mailBody = email.message_from_string(body.decode('utf-8'))
                    print(mailBody)
                    doc.add_paragraph(str(mailBody))
                    doc.save('architecture.docx')

















