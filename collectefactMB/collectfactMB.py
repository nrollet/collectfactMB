import logging
import json
import pprint
import openpyxl
from fetchemail import FetchEmail
from htmlparser import extract_htmltable
from datetime import datetime


logging.basicConfig(
    level=logging.DEBUG, format="%(levelname)s \t %(module)s -- %(message)s"
)
pp = pprint.PrettyPrinter(indent=4)


def myconverter(o):
    if isinstance(o, datetime.datetime):
        return o.__str__()


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


table_list = []
for account in config["ACCOUNTS"].keys():
    logging.info("Accès au compte {}".format(account))
    ID = config["ACCOUNTS"][account]["ID"]
    IMAP_PWD = config["ACCOUNTS"][account]["IMAP_PWD"]
    IMAP_GIS = config["ACCOUNTS"][account]["GIS_PWD"]
    INBOX = config["ACCOUNTS"][account]["INBOX"]
    # print("\n".join([ID,IMAP_PWD, INBOX]))

    mailsrv = FetchEmail(IMAP, ID, IMAP_PWD, INBOX)
    msg_list = mailsrv.fetch_specific_messages(SENDER, SUBJECT)
    # print(msg_list)
    logging.info("messages en attente : {}".format(len(msg_list)))

    for msg in msg_list:
        table = extract_htmltable(msg["body"])
        table_list.append(table)

        for customer in table.keys():
            for invoice in table[customer]:
                ws.append(
                    [
                        datetime.now(),
                        customer,
                        invoice,
                        table[customer][invoice]["montant"],
                    ]
                )

        # parsed = json.loads(table)
        # txt = json.dumps(table, indent=2, sort_keys=True, default=myconverter)
        # with open(account + "_" + msg["num"].decode() + ".json", "w") as f:
        #     f.write(txt)
    # txt = json.dumps(table_list, indent=2, sort_keys=True, default=myconverter)
    # with open("toutjson", "w") as f:
    #     f.write(txt)
    #     print("x-"*20)
    mailsrv.close_connection()

wb.save(XL)
wb.close()

# MAILSRV = origin["SERVER"]
