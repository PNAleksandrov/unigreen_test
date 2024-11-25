import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import urllib3
import re
import os

# Отключаем предупреждение о необходимости верификации
urllib3.disable_warnings()


def create_directory(directory_name):
    # Создаем директорию, если она не существует
    if not os.path.exists(directory_name):
        os.makedirs(directory_name)
    return directory_name

def get_html(date):
    url = f"https://www.atsenergo.ru/nreport?rname=big_nodes_prices_pub&region=eur&rdate={date.strftime('%Y%m%d')}"
    response = requests.get(url, verify=False)
    return response.text

def extract_fid(html):
    soup = BeautifulSoup(html, 'html.parser')
    links = soup.find_all('a', href=True)

    for link in links:
        href = link['href']
        if 'fid=' in href:
            fid = href.split('fid=')[1]
            fid = fid.split('&')[0]
            return fid

    return None

def download_file(fid, date, directory):
    url = f"https://www.atsenergo.ru/nreport?fid={fid}&region=eur"
    filename = f"{directory}/prices_{date.strftime('%Y-%m-%d')}.xlsx"

    response = requests.get(url, verify=False)
    with open(filename, 'wb') as file:
        file.write(response.content)
    print(f"Файл {filename} успешно сохранен")

def download_files(start_date, end_date, download_directory):
    result_df = pd.DataFrame(columns=['date', 'value'])

    current_date = start_date
    while current_date <= end_date:
        html = get_html(current_date)
        fid = extract_fid(html)

        if fid:
            download_file(fid, current_date, download_directory)

            file_path = os.path.join(download_directory, f"prices_{current_date.strftime('%Y-%m-%d')}.xlsx")
            df = pd.read_excel(file_path)
            processed_data = process_data(df, current_date)
            result_df = pd.concat([result_df, processed_data])
        else:
            print(f"Не удалось найти fid для даты {current_date.strftime('%Y-%m-%d')}")

        current_date += timedelta(days=1)

def main():
    start_date = datetime(2024, 9, 2)
    end_date = datetime(2024, 9, 9)
    directory_name = "downloaded_xls"

    download_directory = create_directory(directory_name)

    current_date = start_date
    while current_date <= end_date:
        html = get_html(current_date)
        fid = extract_fid(html)

        if fid:
            download_file(fid, current_date, download_directory)
        else:
            print(f"Не удалось найти fid для даты {current_date.strftime('%Y-%m-%d')}")

        current_date += timedelta(days=1)

if __name__ == "__main__":
    main()