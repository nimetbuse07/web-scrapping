import requests
from bs4 import BeautifulSoup
import xlsxwriter

header = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36"}

workbook = xlsxwriter.Workbook('trendyol.xls')
workbook2 = xlsxwriter.Workbook('mediamarkt.xls')
workbook3 = xlsxwriter.Workbook('n11.xls')
workbook4 = xlsxwriter.Workbook("vatan.xls")


def trendyollistele():
    worksheet = workbook.add_worksheet()

    ozellikListesiTrendyol = []
    trendyolIslemciTipi = []
    trendyolDepolama = []
    trendyolIsletimSistemi = []
    trendyolRam = []
    trendyolEkranBoyutu = []
    trendyolFiyat = []
    trendyolLink = []
    trendyolMarka = []
    trendyolBaslik = []
    j = 0
    k = 1
    m = 2
    n = 4
    b = 5
    anaLink = "https://www.trendyol.com"
    for x in range(1, 13):
        r = requests.get(f"https://www.trendyol.com/sr?q=notebook&qt=notebook&st=notebook&lc=103108&os=1&pi={x}")
        soup = BeautifulSoup(r.content, "lxml")
        notebooks = soup.find_all("div", attrs={"class": "prdct-cntnr-wrppr"})
        for notebook in notebooks:
            notebook_linkleri = notebook.find_all("div", attrs={"class": "with-campaign-view"})
            for i in notebook_linkleri:
                link = i.find("div", attrs={"class": "card-border"})
                link_devam = link.a.get("href")
                fiyat = link.find("div", {"class": "product-down"})
                fiyat_devam = fiyat.find("div", {"class": "price-promotion-container"})
                fiyat_devam2 = fiyat_devam.find("div", {"class": "discountedPriceBox"})
                fiyat_devam3 = fiyat_devam2.find("div", {"class": "prc-box-dscntd"}).text
                trendyolFiyat.append(fiyat_devam3)
                trendyolLink.append(anaLink + link_devam)
                marka = link.find("div", {"class": "product-down"})
                marka2 = marka.find("div", {"class": "prdct-desc-cntnr"})
                marka3 = marka2.find("div", {"class": "two-line-text"})
                marka4 = marka3.find("span", {"class": "prdct-desc-cntnr-ttl"}).text
                trendyolMarka.append(marka4)
                baslik = marka3.text
                trendyolBaslik.append(baslik)

                detay = requests.get(anaLink + link_devam, headers=header)
                detay_soup = BeautifulSoup(detay.content, "html.parser")

                notebookBoard = detay_soup.find('div', {'class': 'starred-attributes'})
                for li in notebookBoard.find_all('li'):
                    ozellik = li.find_all('span')[1].text.strip()
                    ozellikListesiTrendyol.append(ozellik)
                while (j < len(ozellikListesiTrendyol)):
                    trendyolIslemciTipi.append(ozellikListesiTrendyol[j])
                    j += 6
                while (k < len(ozellikListesiTrendyol)):
                    trendyolDepolama.append(ozellikListesiTrendyol[k])
                    k += 6
                while (m < len(ozellikListesiTrendyol)):
                    trendyolIsletimSistemi.append(ozellikListesiTrendyol[m])
                    m += 6
                while (n < len(ozellikListesiTrendyol)):
                    trendyolRam.append(ozellikListesiTrendyol[n])
                    n += 6
                while (b < len(ozellikListesiTrendyol)):
                    trendyolEkranBoyutu.append(ozellikListesiTrendyol[b])
                    b += 6
    # data via the write() method.

    for t in range(1, 251):
        x = str(t)
        worksheet.write(f'A{t}', "" + x + "")
        worksheet.write(f'B{t}', "" + trendyolEkranBoyutu[t - 1] + "")
        worksheet.write(f'C{t}', "" + trendyolDepolama[t - 1] + "")
        worksheet.write(f'D{t}', "" + trendyolIslemciTipi[t - 1] + "")
        worksheet.write(f'E{t}', "" + trendyolIsletimSistemi[t - 1] + "")
        worksheet.write(f'F{t}', "" + trendyolRam[t - 1] + "")
        worksheet.write(f'G{t}', "" + trendyolFiyat[t - 1] + "")
        worksheet.write(f'H{t}', "" + trendyolMarka[t - 1] + "")
        worksheet.write(f'I{t}', "" + trendyolBaslik[t - 1] + "")
        worksheet.write(f'J{t}', "" + trendyolLink[t - 1] + "")
    workbook.close()
    return


def mmlistele():
    worksheet = workbook2.add_worksheet()
    ozellikListesimm = []
    mmRam = []
    mmIslemci = []
    mmDepolama = []
    mmEkran = []
    mmLink = []
    mmFiyat = []
    mmBaslik = []
    k = 0
    l = 1
    m = 3
    n = 4
    ana_link = "https://www.mediamarkt.com.tr"
    for x in range(1, 4):
        r = requests.get(
            f"https://www.mediamarkt.com.tr/tr/category/_laptop-504926.html?searchParams=&sort=suggested&view=&page={x}")
        soup = BeautifulSoup(r.content, "lxml")
        notebooks = soup.find_all("ul", attrs={"class": "products-list"})
        for notebook in notebooks:
            notebook_linkleri = notebook.find_all("div", attrs={"class": "product-wrapper"})
            for i in notebook_linkleri:
                link = i.find("div", attrs={"class": "content"})
                link_devam = link.a.get("href").strip()
                fiyat = i.find("aside", attrs={"class": "alt"})
                fiyat_devam = fiyat.find("div", attrs={"class": "infobox"})
                fiyat_devam2 = fiyat_devam.find("div", {"class": "price-box"}).text.split()[0]
                mmLink.append(ana_link + link_devam)
                mmFiyat.append(fiyat_devam2)
                baslik = link.h2.text.strip()
                mmBaslik.append(baslik)

                detay = requests.get(ana_link + link_devam, headers=header)
                detay_soup = BeautifulSoup(detay.content, "html.parser")

                notebookBoard = detay_soup.find('dl', {'class': 'product-details'})
                for dd in notebookBoard.find_all('dd'):
                    ozellik = dd.text
                    ozellikListesimm.append(ozellik)
                while (k < len(ozellikListesimm)):
                    mmRam.append(ozellikListesimm[k])
                    k += 7
                while (l < len(ozellikListesimm)):
                    mmIslemci.append(ozellikListesimm[l])
                    l += 7
                while (m < len(ozellikListesimm)):
                    mmDepolama.append(ozellikListesimm[m])
                    m += 7
                while (n < len(ozellikListesimm)):
                    mmEkran.append(ozellikListesimm[n])
                    n += 7
    # data via the write() method.
    for t in range(1, 78):
        x = str(t)
        worksheet.write(f'A{t}', "" + x + "")
        worksheet.write(f'B{t}', "" + mmRam[t - 1] + "")
        worksheet.write(f'C{t}', "" + mmIslemci[t - 1] + "")
        worksheet.write(f'D{t}', "" + mmDepolama[t - 1] + "")
        worksheet.write(f'E{t}', "" + mmEkran[t - 1] + "")
        worksheet.write(f'F{t}', "" + mmFiyat[t - 1] + "")
        worksheet.write(f'G{t}', "" + mmBaslik[t - 1] + "")
        worksheet.write(f'H{t}', "" + mmLink[t - 1] + "")
    workbook2.close()
    return


def n11listele():
    worksheet = workbook3.add_worksheet()
    ozellikListesi = []
    n11Islemci = []
    n11Ram = []
    n11Marka = []
    n11Ekran = []
    n11IsletimSistemi = []
    n11Fiyat = []
    n11Link = []
    n11Puan = []
    n11Baslik = []
    k = 0
    l = 1
    m = 2
    n = 3
    b = 4

    for x in range(1, 22):
        r = requests.get(f"https://www.n11.com/arama?q=notebook&ipg={x}")
        soup = BeautifulSoup(r.content, "lxml")
        notebooks = soup.find_all("li", attrs={"class": "column"})
        for notebook in notebooks:
            notebook_linkleri = notebook.find_all("div", attrs={"class": "columnContent"})
            for i in notebook_linkleri:
                link = i.find("div", attrs={"class": "pro"})
                link_devam = link.a.get("href")
                fiyat = link.find("div", attrs={"class": "proDetail"})
                fiyat_devam = fiyat.find("div", attrs={"class": "priceContainer"})
                fiyat_devam2 = fiyat_devam.find("span", attrs={"class": "priceEventClick"}).text.strip()
                n11Link.append(link_devam)
                n11Fiyat.append(fiyat_devam2)
                baslik = link.find("a")
                baslik2 = baslik.find("h3", {"class": "productName"}).text
                n11Baslik.append(baslik2)

                detay = requests.get(link_devam, headers=header)
                detay_soup = BeautifulSoup(detay.content, "html.parser")

                notebookBoard = detay_soup.find('div', {'class': 'unf-attribute-cover'})
                notebookBoard2 = detay_soup.find('div', {'class': 'unf-p-title'})
                notebookBoard3 = notebookBoard2.find('div', {'class': 'proRatingHolder'})

                for di in notebookBoard3.find_all('div', {'class': 'ratingCont'}):
                    puan = di.find('strong').text
                    n11Puan.append(puan)
                for div in notebookBoard.find_all('div'):
                    ozellik = div.find_all('strong')[0].text.strip()
                    ozellikListesi.append(ozellik)
                while (k < len(ozellikListesi)):
                    n11Islemci.append(ozellikListesi[k])
                    k += 5
                while (l < len(ozellikListesi)):
                    n11Ram.append(ozellikListesi[l])
                    l += 5
                while (m < len(ozellikListesi)):
                    n11Marka.append(ozellikListesi[m])
                    m += 5
                while (n < len(ozellikListesi)):
                    n11Ekran.append(ozellikListesi[n])
                    n += 5
                while (b < len(ozellikListesi)):
                    n11IsletimSistemi.append(ozellikListesi[b])
                    b += 5
    # data via the write() method.

    for t in range(1, 501):
        x = str(t)
        worksheet.write(f'A{t}', "" + x + "")
        worksheet.write(f'B{t}', "" + n11Islemci[t - 1] + "")
        worksheet.write(f'C{t}', "" + n11Ram[t - 1] + "")
        worksheet.write(f'D{t}', "" + n11Marka[t - 1] + "")
        worksheet.write(f'E{t}', "" + n11Ekran[t - 1] + "")
        worksheet.write(f'F{t}', "" + n11IsletimSistemi[t - 1] + "")
        worksheet.write(f'G{t}', "" + n11Fiyat[t - 1] + "")
        worksheet.write(f'H{t}', "" + n11Puan[t - 1] + "")
        worksheet.write(f'I{t}', "" + n11Baslik[t - 1] + "")
        worksheet.write(f'J{t}', "" + n11Link[t - 1] + "")
    workbook3.close()
    return


def vatanlistele():
    worksheet = workbook4.add_worksheet()
    ozellikListesiVatan = []
    vatanIslemci = []
    vatanRam = []
    vatanEkran = []
    vatanDepolama = []
    vatanIsletimSistemi = []
    vatanFiyat = []
    vatanLink = []
    vatanPuan = []
    vatanBaslik = []

    k = 3
    l = 9
    m = 16
    n = 25
    b = 50

    ana_link = "https://www.vatanbilgisayar.com/"
    for x in range(1, 19):
        r = requests.get(f"https://www.vatanbilgisayar.com/arama/notebook%20bilgisayar/?page={x}")
        soup = BeautifulSoup(r.content, "lxml")
        notebooks = soup.find_all("div", attrs={"class": "clearfix"})
        for notebook in notebooks:
            notebook_linkleri = notebook.find_all("div", attrs={"class": "product-list--list-page"})
            for i in notebook_linkleri:
                link = i.find("div", attrs={"class": "product-list__content"})
                link_devam = link.find("a", {"class": "product-list__link"}).get("href")
                fiyat = link.find("div", {"class": "product-list__cost"})
                fiyat_devam = fiyat.find("span", {"class": "product-list__price"}).text
                vatanLink.append(ana_link + link_devam)
                vatanFiyat.append(fiyat_devam)
                baslik = link.find("a", {"class": "product-list__link"})
                baslik2 = baslik.find("div", {"class": "product-list__product-name"})
                baslik3 = baslik2.find("h3").text
                vatanBaslik.append(baslik3)

                detay = requests.get(ana_link + link_devam, headers=header)
                detay_soup = BeautifulSoup(detay.content, "lxml")

                notebookBoard = detay_soup.find_all("tr", {"data-count": "0"})
                notebookBoard2 = detay_soup.find("div", class_="col-lg-8 col-md-8 col-sm-8 col-xs-12")
                for q in notebookBoard2.find("strong", attrs={'id': 'averageRankNum'}):
                    puan = q.text
                    vatanPuan.append(puan)
                for c in notebookBoard:
                    ozellik = c.find("p").text
                    ozellikListesiVatan.append(ozellik)
                while (k < len(ozellikListesiVatan)):
                    vatanIslemci.append(ozellikListesiVatan[k])
                    k += 6
                while (l < len(ozellikListesiVatan)):
                    vatanRam.append(ozellikListesiVatan[l])
                    l += 6
                while (m < len(ozellikListesiVatan)):
                    vatanEkran.append(ozellikListesiVatan[m])
                    m += 6
                while (n < len(ozellikListesiVatan)):
                    vatanDepolama.append(ozellikListesiVatan[n])
                    n += 6
                while (b < len(ozellikListesiVatan)):
                    vatanIsletimSistemi.append(ozellikListesiVatan[b])
                    b += 6
    # data via the write() method.

    for t in range(1, 291):
        x = str(t)
        worksheet.write(f'A{t}', "" + x + "")
        worksheet.write(f'B{t}', "" + vatanIslemci[t - 1] + "")
        worksheet.write(f'C{t}', "" + vatanRam[t - 1] + "")
        worksheet.write(f'D{t}', "" + vatanEkran[t - 1] + "")
        worksheet.write(f'E{t}', "" + vatanDepolama[t - 1] + "")
        worksheet.write(f'F{t}', "" + vatanIsletimSistemi[t - 1] + "")
        worksheet.write(f'G{t}', "" + vatanFiyat[t - 1] + "")
        worksheet.write(f'H{t}', "" + vatanPuan[t - 1] + "")
        worksheet.write(f'I{t}', "" + vatanBaslik[t - 1] + "")
        worksheet.write(f'J{t}', "" + vatanLink[t - 1] + "")

    workbook4.close()
    return


trendyollistele()
mmlistele()
n11listele()
vatanlistele()
