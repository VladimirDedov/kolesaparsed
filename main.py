import requests, openpyxl, lxml, datetime, os, time, shutil
from fake_useragent import UserAgent
from bs4 import BeautifulSoup

def collect_data():
    list_name, list_price, list_links, list_year=[],[],[],[]
    ua = UserAgent()
    headers={"user-agent":f"{ua.random}"}
    url = f"https://kolesa.kz/cars/region-severokazakhstanskaya-oblast/?auto-car-transm=2&auto-sweel=1&year[from]=2000&price[to]=2500000"
    responce = requests.get(url=url, headers=headers)
    soup = BeautifulSoup(responce.text, "lxml")
    numbers_page= soup.find("div", class_="pager").find_all("li")
    count = int(numbers_page[-1].find("span").text)

    for i in range(1, count+1):
        responce = requests.get(
            url=f"https://kolesa.kz/cars/region-severokazakhstanskaya-oblast/?auto-car-transm=2&auto-sweel=1&year[from]=2000&price[to]=2500000&page={i}",
            headers=headers)
        soup = BeautifulSoup(responce.text, "lxml")
        cars = soup.find_all("div", class_="a-card__info")  # список всех карточек со страницы
        for items in cars:
            list_name.append(items.find("a", class_="a-card__link").text.strip())
            list_year.append(items.find("p", class_="a-card__description").text.strip())
            price=items.find("span", class_="a-card__price").text
            price=''.join(price.split())
            list_price.append(price)
            list_links.append("https://kolesa.kz"+items.find("a", class_="a-card__link").get("href"))
        time.sleep(1)#protect the server

    write_XLS(list_name, list_price, list_links,list_year)
    answer=input("Download foto? y/n - ")
    if answer == 'y':
        download_images(list_links, list_name, headers)

def write_XLS(list_name, list_price, list_links, list_year):
    book = openpyxl.open("car.xlsx")
    if str(datetime.date.today()) not in book.sheetnames:
        book.create_sheet(f"{str(datetime.date.today())}")
    sheet = book[f"{str(datetime.date.today())}"]
    count_line = 1

    for column in range(1, len(list_name) + 1):
        sheet[f'A{column}'] = list_name[column - 1]
        sheet[f'B{column}'] = list_price[column - 1]
        sheet[f'C{column}'] = list_year[column - 1]
        sheet[f'D{column}'] = list_links[column - 1]

    book.save("car.xlsx")
    book.close()
    print("All cars are added to the file car.xlsx. Successfully recorded!")

def download_images(list_link_car, list_name, headers):
    i=0
    if os.path.exists(r"C:\Pyton\Parsing\Kolesa\FotoCar"):
        shutil.rmtree(r"C:\Pyton\Parsing\Kolesa\FotoCar")#remove all from FotoСar
        os.mkdir("FotoCar/")
    else:
        os.mkdir("FotoCar/")
    for url in list_link_car:
        time.sleep(3)
        print(url)
        responce = requests.get(url, headers=headers)
        soup=BeautifulSoup(responce.text, "lxml")

        # search link big foto
        link = soup.find("div", class_="offer__gallery gallery").find("img").get("src")
        img_data = requests.get(link, verify=False).content

        #Create directory and download foto
        if os.path.exists("FotoCar/"+f"{list_name[i]}"):
            os.mkdir("FotoCar/" + f"{list_name[i]}" + f"-{i}")
            print("Download main foto")
            with open("FotoCar/" + f"{list_name[i]}"+f"-{i}/" + f"{list_name[i]}.jpg", "wb") as handler:
                handler.write((img_data))
        else:
            os.mkdir("FotoCar/"+f"{list_name[i]}")
            print("Download main foto")
            with open("FotoCar/" + f"{list_name[i]}/" + f"{list_name[i]}.jpg", "wb") as handler:
                handler.write((img_data))
        i+=1
    print("All photos have been successfully downloaded to the folder FotoCar!")
def main():
   collect_data()

if __name__ == "__main__":
    main()
