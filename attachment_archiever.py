import datetime as dt
import email
import imaplib
import json
import os

from tqdm import tqdm

if __name__ == '__main__':
    # load config with field: [server, username, password, storage_root]
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    m = imaplib.IMAP4_SSL(config['server'])
    m.login(config['username'], config['password'])
    m.select()

    resp, items = m.search(None, "(ALL)")

    emails = items[0].split()
    # email_id = emails[0]
    with tqdm(emails) as pbar:
        for email_id in emails:
            resp, data = m.fetch(email_id, "(RFC822)")
            email_body = data[0][1]
            mail = email.message_from_string(email_body.decode('gb2312'))

            sender = mail['FROM']
            if ' ' in sender:
                sender = sender.split(' ')[1]
            sender = sender.replace('<', '')
            sender = sender.replace('>', '')
            if '?' in sender:
                sender = email.header.decode_header(sender)[-1][0].decode()
            folder_name = os.path.join(config['storage_root'], sender)

            for part in mail.walk():
                if part.get_content_maintype() != 'multipart':
                    file_name = part.get_filename()
                    if file_name:
                        os.makedirs(folder_name, exist_ok=True)

                        raw_date_str = mail['Date'][:31]
                        if raw_date_str[-1] == ' ':
                            raw_date_str = raw_date_str[:-1]
                        if raw_date_str[:3].isalpha():
                            raw_date_str = raw_date_str[5:]
                        time = dt.datetime.strptime(raw_date_str, '%d %b %Y %H:%M:%S %z')
                        time_str = time.strftime('%Y%m%d%H%M%S')

                        if '?' in file_name:
                            content, charset = email.header.decode_header(file_name)[0]
                            file_name = content.decode(charset)
                        full_path = os.path.join(folder_name, file_name)
                        if os.path.isfile(full_path):
                            file_name = '_'.join([time_str, file_name])
                            full_path = os.path.join(folder_name, file_name)
                        print(email_id, file_name)

                        with open(full_path, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        m.store(email_id, '+FLAGS', r'\Seen')
            pbar.update()
    m.close()
    m.logout()
