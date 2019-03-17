import openpyxl
import logging
from datetime import datetime


def update_excel(
    filename,
    date_collecte=datetime(1800, 1, 1),
    compte_msg="",
    client="",
    date_facture=datetime(1800, 1, 1),
    numero="",
    montant=0.0,
    url="",
):

    try:
        wb = openpyxl.load_workbook(filename=filename)
    except OSError as e:
        logging.error("{}".format(e))
        wb = openpyxl.Workbook()

    ws = wb.active
    row = [date_collecte, compte_msg, client, date_facture, numero, montant, url]
    ws.append(row)
    wb.save(filename)
    wb.close()


for x in range(10):
    update_excel(
        "toto.xlsx",
        compte_msg="ABC",
        client="JUSTOM",
        numero="12345",
        date_collecte=datetime(2012, 12, 12),
        date_facture=datetime(2013, 12, 12),
    )
