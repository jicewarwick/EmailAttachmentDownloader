#! /usr/bin/python
import datetime as dt
import email
import imaplib
import json
import os

from tqdm import tqdm


class EmailAttachmentDownloader(object):
    def __init__(self, server: str, email_address: str, password: str):
        super().__init__()
        self.server = server
        self.email_address = email_address
        self.password = password
        self.m = None

    def login(self):
        self.m = imaplib.IMAP4_SSL(self.server)
        self.m.login(self.email_address, self.password)
        self.m.select()

    def logout(self):
        self.m.close()
        self.m.logout()

    def __enter__(self):
        self.login()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.logout()

    @staticmethod
    def parse_sender(sender: str) -> str:
        if ' ' in sender:
            sender = sender.split(' ')[1]
        sender = sender.replace('<', '')
        sender = sender.replace('>', '')
        if '?' in sender:
            sender = email.header.decode_header(sender)[-1][0].decode()
        return sender

    @staticmethod
    def parse_datetime(email_date: str) -> dt.datetime:
        raw_date_str = email_date[:31]
        if raw_date_str[-1] == ' ':
            raw_date_str = raw_date_str[:-1]
        if raw_date_str[:3].isalpha():
            raw_date_str = raw_date_str[5:]
        time = dt.datetime.strptime(raw_date_str, '%d %b %Y %H:%M:%S %z')
        return time

    @staticmethod
    def parse_file_name(file_name: str) -> str:
        if '?' in file_name:
            content, charset = email.header.decode_header(file_name)[0]
            file_name = content.decode(charset)

        return file_name

    def download_all_attachments(self, output_dir: str):
        resp, items = self.m.search(None, "(ALL)")
        emails = items[0].split()
        # email_id = emails[27]
        with tqdm(emails) as progress_bar:
            for email_id in emails:
                resp, data = self.m.fetch(email_id, "(RFC822)")
                email_body = data[0][1]
                mail = email.message_from_bytes(email_body)

                sender = self.parse_sender(mail['FROM'])
                folder_name = os.path.join(output_dir, sender)
                sent_datetime = self.parse_datetime(mail['Date'])

                for part in mail.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue

                    file_name = part.get_filename()
                    if file_name is None:
                        continue

                    file_name = self.parse_file_name(file_name)
                    os.makedirs(folder_name, exist_ok=True)
                    full_path = os.path.join(folder_name, file_name)

                    if os.path.isfile(full_path):
                        time_str = sent_datetime.strftime('%Y%m%d%H%M%S')
                        file_name = '_'.join([time_str, file_name])
                    full_path = os.path.join(folder_name, file_name)
                    print(email_id, file_name)

                    with open(full_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))

                    self.m.store(email_id, '+FLAGS', r'\Seen')

                progress_bar.update()


if __name__ == '__main__':
    # load config with field: [server, username, password, storage_root]
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    attachment_downloader = EmailAttachmentDownloader(config['server'], config['username'], config['password'])
    attachment_downloader.login()

    self = attachment_downloader

    attachment_downloader.download_all_attachments('.')
