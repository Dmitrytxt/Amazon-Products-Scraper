from selenium import webdriver
from bs4 import BeautifulSoup
from spreadsheet import Spreadsheet
import googleapiclient.errors
import csv


def create_driver(driver_path):
    driver = webdriver.Chrome(executable_path=driver_path)
    return driver


def get_url(search_name):
    base_url = 'https://www.amazon.com/s?k={}'
    search_name = search_name.replace(' ', '+')
    full_url = base_url.format(search_name)
    full_url += '&page={}'
    return full_url


def get_item(item):
    try:
        price = item.find('span', {'class': 'a-offscreen'}).text
    except AttributeError:
        return
    try:
        rating = item.i.text
    except AttributeError:
        rating = ''
    try:
        reviews = item.find('span', {'class': 'a-size-base', 'dir': 'auto'}).text
    except AttributeError:
        reviews = ''
    name = item.h2.a.text.strip()
    link = 'https://www.amazon.com' + item.h2.a.get('href')

    result = [name, price, rating, reviews, link]
    return result


def import_to_googlesheets(records, search_term, credentials):
    ss = Spreadsheet(credentials, debug_mode=False)
    try:
        f = open('temp.txt', 'x')
        email = input('Email: ')
        ss.create('Amazon', search_term)
        ss.share_with_email_for_writing(email)
        f.write(ss.spreadsheet_id)
        f.close()
    except FileExistsError:
        f = open('temp.txt', 'r')
        ss.set_spreadsheet_by_id(f.read())
        try:
            ss.add_sheet(search_term)
        except googleapiclient.errors.HttpError:
            ss.set_sheet_by_title(search_term)
            ss.clear_sheet()
        f.close()
    finally:
        ss.prepare_set_values('A1:E1', [['Name', 'Price', 'Rating', 'Reviews', 'Url']])
        ss.prepare_set_values('A2:E%d' % (len(records) + 1), records)
        ss.run_prepared()


def save_to_scv(name, records):
    header = ['Name', 'Price', 'Rating', 'Reviews', 'Url']
    with open(name + '.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(records)


def main(search_term, credentials, driver_path):
    url = get_url(search_term)
    driver = create_driver(driver_path)
    records = []

    for page in range(1, 21):
        driver.get(url.format(page))
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        results = soup.find_all('div', {'data-component-type': 's-search-result'})

        for i in results:
            record = get_item(i)
            if record:
                records.append(record)

    import_to_googlesheets(records, search_term, credentials)
    save_to_scv(search_term, records)

    driver.close()


if __name__ == '__main__':
    term = input('Search term: ')
    CREDENTIALS = 'creds.json'
    d_path = 'C:\chromedriver\chromedriver'
    main(term, CREDENTIALS, d_path)
