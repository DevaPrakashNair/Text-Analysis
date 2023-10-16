import requests
from bs4 import BeautifulSoup

def text_extraction(d):
    for id, url in d.items():
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        a = soup.find_all("div", {"class": "td-post-content tagdiv-type"})
        text = soup.find("title").get_text()
        path = r"C:\Users\Deva\Study_material\Blackcoffer\20211030 Test Assignment\Py\text"+"/"+str(id)+".txt"
        for i in a:
            text += i.get_text()
        file = open(path, "w", encoding="utf-8")
        file.write(text)

    return "Text extracted"