import logging
import bs4 as BeautifulSoup
from datetime import datetime

def extract_htmltable(html):
    soup = BeautifulSoup.BeautifulSoup(html, features="html.parser")
    table = soup.find_all("tr", {"class": "texteNormal"})
    rows = []
    for item in table:
        ttr = item.find_all("td", {"class": "celluleInterne"})
        cols = []
        for tr in ttr:
            cols.append(tr.get_text())
            aa = tr.find_all("a")
            for a in aa:
                cols.append(a["href"])
        rows.append(cols)

    main_dic = {}

    for row in rows:
        # if "MARTIN-BROWER FRANCE BEAUVAIS" in row and len(row) == 9:
        if row[0].startswith("MARTIN-BROWER FRANCE") and len(row) == 9:

            customer = row[1]
            invoice = row[2]
            date = ""
            amount = 0.0
            try:
                date = datetime.strptime(row[3], "%m/%d/%Y")
            except ValueError:
                logging.error("date illisible : {}\n{}".format(row[3], "_".join(row)))
            try:
                amount = float(row[4])
            except ValueError:
                logging.error("montant illisible : {}\n{}".format(row[4], "_".join(row)))
            url_pdf = row[6]
            url_edi = row[8]


            main_dic.setdefault(customer, {})
            main_dic[customer].setdefault(invoice, {})
            main_dic[customer][invoice].update(
                {
                    "date": date,
                    "montant": amount,
                    "url_pdf": url_pdf,
                    "url_edi": url_edi
                }    
            )
        elif "TOUS.PDF" in row or "TOUS.EDI" in row :
            pass
        else:
            logging.warning("cette ligne a été rejetée : {}".format("_".join(row)))

    return main_dic



if __name__ == "__main__":
    import pprint
    import logging
    logging.basicConfig(
        level=logging.DEBUG, format="%(module)s \t %(levelname)s -- %(message)s"
    )    

    pp = pprint.PrettyPrinter(indent=4)

    with open("tests/body.html", "r") as f:
        data = f.read()

    pp.pprint(extract_htmltable(data))

