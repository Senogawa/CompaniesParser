import pandas as pd
import requests
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from selenium import webdriver
import time
import re
import regex
import csv
from pyexcel.cookbook import merge_all_to_a_book
import glob
import os
import openpyxl as pxl

def write_in_worksheet(work_sheet_name: str):

    df_not_redacted = pd.read_excel("output_not_redacted.xlsx")
    df_data = {
        "Компания": df_not_redacted["Компания"].tolist(),
        "Телефон": df_not_redacted["Телефон"].tolist(),
        "Почта": df_not_redacted["Почта"].tolist(),
        "Адрес": df_not_redacted["Адрес"].tolist()
    }

    df = pd.DataFrame({
        "Компания": df_data["Компания"],
        "Телефон": df_data["Телефон"],
        "Почта": df_data["Почта"],
        "Адрес": df_data["Адрес"]
    })
    
    excel_book = pxl.load_workbook("output.xlsx")
    with pd.ExcelWriter("output.xlsx", "openpyxl") as writer:
        writer.book = excel_book
        print(df)
        df.to_excel(writer, work_sheet_name, index = False)
        writer.save()


def read_exel(lists: int):
    companies = list()
    for list_name in range(1, lists + 1):
        df = pd.read_excel("data.xlsx",sheet_name = str(list_name), index_col = 0)
        companies.append(df["Unnamed: 1"].tolist())
        #print(df)
    return companies


def test_write(text):
    with open("index.html", "w") as f:
        f.write(text)


def main(lists, headless_or_not):
    df = pd.DataFrame()
    df.to_excel("output.xlsx")
    companies = read_exel(lists)
    #companies = [["1","2","3"],["1","2","3"],["1","2","3"]]
    for sheet in range(1, len(companies) + 1):

        with open("data_redacted.csv", "w") as f:
            writer = csv.writer(f)
            writer.writerow(
                ["Компания" ,"Телефон", "Почта", "Адрес"]
            )
        for company in companies[sheet - 1]:
            while True:
                chrome_options = webdriver.ChromeOptions()
                if headless_or_not:
                    chrome_options.headless = True

                #chrome_options.add_argument(f"user-agent={UserAgent().random}")
                chrome_options.add_argument("--disable-blink-features=AutomationControlled")
                driver = webdriver.Chrome(executable_path = "./chromedriver", options = chrome_options)
                try:
                    driver.get(f"https://duckduckgo.com/?q={company}&t=h_&ia=web")
                except:
                    print("BROWSER")
                    driver.quit()
                    continue
                time.sleep(3)
                #test_write(driver.page_source)
                soup = BeautifulSoup(driver.page_source, "lxml")
                driver.quit()
                try:
                    table = soup.find("div", attrs = {"class":"results js-results"}).find_all("div", attrs = {"class":"nrn-react-div"})
                    break
                except AttributeError:
                    print("ATRIBUTE ERROR!")
                    time.sleep(2)
                    continue
            table = table[:5]
            links = [link.find("a", attrs = {"data-testid":"result-title-a"}).attrs["href"] for link in table]

            numbers = ""
            emails = ""
            adressess = ""

            for link in links:
                headers = {
                    "user-agent":UserAgent().random
                }
                try:
                    res = requests.get(link, headers = headers, timeout = 15)

                except:
                    print("!")
                    continue

                text = BeautifulSoup(res.text, "lxml").text
                text = text.strip()

                number = regex.findall(r"(?<=^|\s|>|\;|\:|\))(?:\+|7|8|9|\()[\d\-\(\) ]{8,}\d", text)
                email = regex.findall(r"([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})", text)
                adress = regex.findall(r"\d+[ ](?:[A-Za-z0-9.-]+[ ]?)+(?:Avenue|Lane|Road|Boulevard|Drive|Street|Ave|Dr|Rd|Blvd|Ln|St)\.?", text)

                number = ", ".join(number)
                email = ", ".join(email)
                adress = ", ".join(adress)

                numbers += f", {number}".strip()
                emails += f", {email}".strip()
                adressess += f", {adress}".strip()

                print(numbers)
                print(emails)
                print(adressess)

            #добавление в csv
            with open("data_redacted.csv", "a") as f:
                writer = csv.writer(f)
                writer.writerow(
                    [company, numbers, emails, adressess]
                )
            
            merge_all_to_a_book(glob.glob("./data_redacted.csv"), "output_not_redacted.xlsx")
            write_in_worksheet(str(sheet))
if __name__ == "__main__":
    while True:
        lists = input("Введите количество листов в документе\nИмя листа должно совпадать с его номером: ")
        if lists.isdigit():
            lists = int(lists)
            break
        else:
            print("Вводите только цифры")
    while True:
        answer = input("Скрыть браузер при работе?\nY или N: ")
        if answer == "Y":
            headless = True
            break

        elif answer == "N":
            headless = False
            break

        else:
            print("Введите Y или N")
            continue

    main(lists, headless)