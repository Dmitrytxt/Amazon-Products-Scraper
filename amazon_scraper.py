from selenium import webdriver
from bs4 import BeautifulSoup


def create_driver():
    driver = webdriver.Chrome(executable_path='C:\chromedriver\chromedriver')
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


def main(search_term):
    url = get_url(search_term)
    driver = create_driver()
    records = []

    for page in range(1, 21):
        driver.get(url.format(page))
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        results = soup.find_all('div', {'data-component-type': 's-search-result'})

        for i in results:
            record = get_item(i)
            if record:
                records.append(record)

    driver.close()

    for i in records:
        print(i)


if __name__ == '__main__':
    term = input('Search term: ')
    main(term)
