import email
import imaplib
import os
import logging
import json
import random
import string
from datetime import datetime


class FetchEmail:

    connection = None
    error = None

    def __init__(self, mail_server, username, password, folder):
        """
        Init IMAP connection
        params : mail server, username, passw
        """
        self.connection = imaplib.IMAP4_SSL(mail_server)
        status, _ = self.connection.login(username, password)
        logging.info("connection : " + status)
        dat = self.connection.select(
            '"{}"'.format(folder), readonly=False
        )  # so we can mark mails as read
        logging.debug('accès dossier : "{}", {}'.format(folder, dat))

    def close_connection(self):
        """
        Close the connection to the IMAP server
        """
        self.connection.close()

    def save_attachment_old(self, msg, prefix, download_folder):
        """
        Given a message, save its attachments to the specified
        download folder (default is /tmp)

        return: file path to attachment
        """
        att_path = "No attachment found."
        att_list = []
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            filename = prefix + "_" + part.get_filename()
            if not os.path.isdir(download_folder):
                os.mkdir(download_folder)

            att_path = os.path.join(download_folder, filename)

            fp = open(att_path, "wb")
            fp.write(part.get_payload(decode=True))
            fp.close()
            att_list.append(att_path)
        return att_list

    def archive_message(self, uid, dest):
        """
        Archive le message dans un sous-dossier
        """
        dest = '"{}"'.format(dest)
        logging.debug("archivage de {} vers {}".format(uid, dest))
        self.connection.create(dest)
        self.connection.uid("COPY", uid, dest)
        self.connection.uid("STORE", uid, "+FLAGS", "\\Deleted")
        self.connection.expunge()

    def fetch_unread_messages(self):
        """
        Retrieve unread messages
        """
        emails = []
        result, messages = self.connection.search(None, "UnSeen")
        if result == "OK":
            for message in messages[0].decode("utf-8").split(" "):
                try:
                    _, data = self.connection.fetch(message, "(RFC822)")
                except:
                    print("No new emails to read.")
                    self.close_connection()
                    exit()

                msg = email.message_from_bytes(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append(msg)
                _, data = self.connection.store(message, "+FLAGS", "\\Seen")

            return emails

        self.error = "Failed to retreive emails."
        return emails

    def fetch_all_messages(self):
        """
        Retrieve all messages in folder
        Returns a list with message uid and
        message object instance
        """
        emails = []
        result, data = self.connection.uid("search", None, "ALL")
        uid_list = email.message_from_bytes(data[0]).__str__().split()
        if result == "OK":
            for uid in uid_list:
                try:
                    result, data = self.connection.uid("fetch", uid, "(RFC822)")
                except:
                    print("No emails to read.")
                    self.close_connection()
                    exit()

                msg = email.message_from_bytes(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append([uid, msg])

        return emails

    def fetch_messages_from(self, sender):
        """
        Recupère message de l'expéditeur
        Retourn un dict :
        - clé : numero du message
        - valeur : instance email.message
        """
        emails = []
        # emails = {}
        result, messages = self.connection.search(None, "FROM", sender)
        if result == "OK":
            for message in messages[0].decode("utf-8").split(" "):
                try:
                    _, data = self.connection.fetch(message, "(RFC822)")
                except:
                    print("No emails to read.")
                    self.close_connection()
                    exit()

                msg = email.message_from_bytes(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append([message, msg])

            return emails

        self.error = "Failed to retreive emails."
        return emails

    def fetch_specific_messages(self, sender, subject):

        messages = []
        # mbox_response, msgnums = self.connection.search(None, 'FROM', sender)
        sender = ""'(FROM "{}")'"".format(sender)
        mbox_response, msgnums = self.connection.uid("search", None, sender)
        self.connection.uid("search", None, "ALL")
        print("###", msgnums)
        if mbox_response == 'OK':
            for num in msgnums[0].split():
                # retval, rawmsg = self.connection.fetch(num, '(RFC822)')
                retval, rawmsg = self.connection.uid("fetch", num, "(RFC822)")
                if retval != 'OK':
                    logging.error('ERROR getting message n.{}'.format(num))                    
                    continue
                msg = email.message_from_bytes(rawmsg[0][1])
                if subject in msg["Subject"]:
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            typ = part.get_content_type()
                            disp = str(part.get("Content-Disposition"))
                            if typ == 'text/plain' and 'attachment' not in disp:
                                charset == part.get_content_charset()
                                body = part.get_payload(decode=True).decode(encoding=charset, errors="ignore")
                                break
                            elif typ == 'text/html' and 'attachment' not in disp:                                
                                charset = part.get_content_charset()
                                body = part.get_payload(decode=True).decode(encoding=charset, errors="ignore")
                                break
                    else:
                        print ("###", f"{num} is NOT multipart")                        
                        charset = msg.get_content_charset()
                        body = msg.get_payload(decode=True).decode(encoding=charset, errors="ignore")                        
                    messages.append({"num": num, "body": body})

        else:
            logging.warning("no message from {}".format(sender))

        return messages


    def parse_email_address(self, email_address):
        """
        Helper function to parse out the email address from the message
        return: tuple (name, address). Eg. ('John Doe', 'jdoe@example.com')
        """
        return email.utils.parseaddr(email_address)





    

