import email
import imaplib
import os

class FetchEmail():

    connection = None
    error = None

    def __init__(self, settings):
        self.connection = imaplib.IMAP4_SSL(settings.MAIL_SERVER)
        self.connection.login(settings.USERNAME, settings.PASSWORD)
        self.connection.select(readonly=False)

    def fetch_unread_messages(self):
        """
        Retrieve unread messages
        """
        emails = []
        (result, messages) = self.connection.search(None, 'UnSeen')
        if result == "OK":
            for message in messages[0].split(' '):
                try:
                    ret, data = self.connection.fetch(message,'(RFC822)')
                except:
                    print "No new emails to read."
                    self.close_connection()
                    exit()

                msg = email.message_from_string(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append(msg)
                # response, data = self.connection.store(message, '+FLAGS','\\Seen')

            return emails

        self.error = "Failed to retreive emails."
        return emails

    def has_attachment(self, msg):
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            return True
        return False

    def save_attachment(self, msg, download_folder="/tmp"):
        """
        Given a message, save its attachments to the specified
        download folder (default is /tmp)

        return: file path to attachment
        """
        att_path = "No attachment found."
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            att_path = os.path.join(download_folder, filename)

            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        return att_path

    def close_connection(self):
        """
        Close the connection to the IMAP server
        """
        self.connection.close()

if __name__ == '__main__':

    import settings

    fetcher = FetchEmail(settings)

    emails = fetcher.fetch_unread_messages()

    for email in emails:
        if fetcher.has_attachment(email):
            fetcher.save_attachment(email, settings.SAVE_DIRECTORY)

    fetcher.close_connection()
