from os import mkdir, chdir, path, startfile, getcwd
import sys
import requests
from bs4 import BeautifulSoup
import urllib.request
import re
import openpyxl


def leave():
    """ Exits programm """
    return sys.exit()


class Webscrapp:
    def __init__(self, dir_name, n, url, urls):
        self.n = 0
        self.url = "https://www.cv.lv/darba-sludinajumi/informacijas-tehnologijas?page="
        self.dir_name = dir_name
        self.urls = urls
        try:
            chdir(path.join(self.dir_name))
            dir_count = 1
        except FileNotFoundError:
            dir_count = 0
        try:
            if dir_count == 0:
                mkdir(self.dir_name)
                chdir(path.join(self.dir_name))
            else:
                pass
        except FileExistsError:
            pass

    @staticmethod
    def create_file():
        """Creates file if it does not exist"""
        headers = ['Kompānija', 'Amats', 'Alga', 'Vairāk info']
        workbook_name = 'jobs_IT.xlsx'
        wb = openpyxl.Workbook()
        page = wb.active
        page.title = 'darbi IT sfērā'
        page.append(headers)
        wb.save(filename=workbook_name)
        wb.close()

    @staticmethod
    def data_store(dataset):
        """Stores info into excel file"""
        workbook_name = 'jobs_IT.xlsx'
        wb = openpyxl.load_workbook(workbook_name)
        page = wb.active
        details = [dataset[0], dataset[1], dataset[2], dataset[3]]
        page.append(details)
        wb.save(filename=workbook_name)

    def scrapper(self):
        """Scraps information from one of most popular Latvian recruiting sites"""
        parser = 'html.parser'
        self.url = f'{self.url}{self.n}'

        try:
            r = requests.get(self.url)
            while r.status_code < 400:
                resp = urllib.request.urlopen(self.url)
                soup = BeautifulSoup(resp, parser)
                for link in soup.find_all('a', href=True):
                    if link not in self.urls:
                        self.urls.append(link['href'])

                url_string = str(self.url)
                self.url = url_string.split('=', 1)[0]
                self.n += 1
                self.url = f'{self.url}={self.n}'
        except:
            self.ini_list()

    def ini_list(self):
        """Creates an initial list of urls to be processed"""
        text = self.urls
        darbi = []
        for item in text:
            if 'htm' in item:
                darbi.append(item)
        darbi = sorted(darbi)

        prefix_ = 'https:'
        final_list = []
        for x in darbi:
            final_list.append(prefix_ + x)
        self.urls = final_list
        self.parser()

    def parser(self):
        """Parses HTML code into data """
        parser = 'html.parser'
        unproceed = []  # for control of results
        linki = self.urls
        m = 0
        print(f'Piedāvājumu skaits: {len(linki)}')
        if path.isfile('jobs_IT.xlsx'):
            pass
        else:
            self.create_file()
        while m < len(linki):
            try:
                resp = urllib.request.urlopen(linki[m])
                soup = BeautifulSoup(resp, parser)
                company = str(soup.find('meta', property="og:title")).split(',')[-1].split('"')[0].split("'")[0]
                if "<meta" in company:
                    company = "N/A"
                elif "property=" in company:
                    company = company.split('property')[0]
                print("Vakance:")
                position = str(soup.find('title')).split(",")[0].split('- ', maxsplit=1)[-1]
                if '</' in position:
                    position = "N/A"
                try:
                    salary = str(soup(text=re.compile("Alga mēnesī"))).split(': ')[1].split("'")[0]
                except:
                    salary = 'N/A'

                data = (f'Kompānija:\t{company}\nAmats:\t{position}\nAlga mēnesī:\t{salary}\n'
                        f'Vairāk info:\t{linki[m]}\n ')
                print(data)
                dataset = [company, position, salary, linki[m]]

                self.data_store(dataset)
                m += 1
            except:
                print("Different structure")
                try:
                    unproceed.append(linki[m])
                    m += 1
                except IndexError:
                    leave()
        a = getcwd()
        startfile(f'{a}\\jobs_IT.xlsx')
        print(f'Unprocessed: {len(unproceed)}')
        print(unproceed)


def main():
    arg = 'dir1'
    n = 0
    url = ''
    urls = []
    wbs = Webscrapp(arg, n, url, urls)
    wbs.scrapper()


if __name__ == '__main__':
    main()
