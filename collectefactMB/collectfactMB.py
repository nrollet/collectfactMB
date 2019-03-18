import logging
import json
import pprint
import openpyxl
import requests
import os
from fetchemail import FetchEmail
from htmlparser import extract_htmltable
from datetime import datetime


logging.basicConfig(
    level=logging.INFO, format="%(levelname)s \t %(module)s -- %(message)s"
)
pp = pprint.PrettyPrinter(indent=4)


def myconverter(o):
    if isinstance(o, datetime.datetime):
        return o.__str__()


### CHECK FOLDER ###
if not os.path.isdir("doc"):
    os.mkdir("doc")
if not os.path.isdir("log"):
    os.mkdir("log")

### LOAD GLOBAL PARAMS ###
with open("config.json", "r") as f:
    config = json.load(f)

GIS_SRV = config["GLOBAL"]["GIS"]
IMAP = config["GLOBAL"]["IMAP"]
SENDER = config["GLOBAL"]["SENDER"]
SUBJECT = config["GLOBAL"]["SUBJECT"]
XL = config["GLOBAL"]["XLLOGGER"]

### EXCEL LOGGER ####
try:
    wb = openpyxl.load_workbook(filename=XL)
except OSError as e:
    logging.error("{}".format(e))
    wb = openpyxl.Workbook()
ws = wb.active
ws.append([])

## Main loop ####
for account in config["ACCOUNTS"].keys():
    logging.info("Accès au compte {}".format(account))
    ID = config["ACCOUNTS"][account]["ID"]
    IMAP_PWD = config["ACCOUNTS"][account]["IMAP_PWD"]
    GIS_PWD = config["ACCOUNTS"][account]["GIS_PWD"]
    INBOX = config["ACCOUNTS"][account]["INBOX"]

    ### HTTP SESSIONS ###
    logging.info("opening http session : {}".format(GIS_SRV))
    s = requests.session()
    html_payload = {"j_username": ID, "j_password": GIS_PWD}
    try:
        r = s.post(GIS_SRV, data=html_payload)
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        logging.error(e)
        continue
    # Verif contenu du cookie pour authent.
    if not r.cookies.get_dict()["decorator"] == "martinbrower":
        logging.error("Authentication has failed")
        continue

    logging.info("opening imap sessions")
    mailsrv = FetchEmail(IMAP, ID, IMAP_PWD, INBOX)
    msg_list = mailsrv.fetch_specific_messages(SENDER, SUBJECT)
    logging.info("messages en attente : {}".format(len(msg_list)))

    for msg in msg_list:
        table = extract_htmltable(msg["body"])
        for customer in table.keys():
            logging.info(
                "msg {} has {} invoices".format(msg["num"], len(table[customer]))
            )
            for invoice in table[customer]:
                ws.append(
                    [
                        datetime.now(),
                        account,
                        customer,
                        invoice,
                        table[customer][invoice]["montant"],
                        table[customer][invoice]["date"],
                        table[customer][invoice]["url_pdf"],
                        table[customer][invoice]["url_edi"],
                    ]
                )
                # PDF
                try:
                    r = s.get(table[customer][invoice]["url_pdf"])
                    r.raise_for_status()
                except requests.exceptions.HTTPError as e:
                    logging.error(e)
                    continue
                except requests.exceptions.RequestException as e:
                    logging.error(e)
                    continue
                open("doc/" + invoice + ".pdf", "wb").write(r.content)
                # EDI
                try:
                    r = s.get(table[customer][invoice]["url_edi"])
                    r.raise_for_status()
                except requests.exceptions.HTTPError as e:
                    logging.error(e)
                    continue
                except requests.exceptions.RequestException as e:
                    logging.error(e)
                    continue
                open("doc/" + invoice + ".edi", "wb").write(r.content)

        mailsrv.archive_message(msg["num"], INBOX + "/archives")
        
    mailsrv.close_connection()
    s.close()

wb.save(XL)
wb.close()
