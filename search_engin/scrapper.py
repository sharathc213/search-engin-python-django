
import requests
from bs4 import BeautifulSoup
from bs4 import BeautifulSoup
import numpy as np
from time import sleep
from random import randint
from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import re
driver = webdriver.Chrome(executable_path="./chromedriver")
person=[]
books=[]
database=[]
sl_no = 1
page="https://pureportal.coventry.ac.uk/en/persons?format=&organisationIds=e5cd4a48-01f9-43af-87fb-74738099f357&nofollow=true&page=0"
driver.get(page)  
soup = BeautifulSoup(driver.page_source, 'html.parser')
pages =soup.find_all("li", {"class": "search-pager-information"})
page_no = re.findall('\d+',str(pages[0]))
content = soup.find_all("div", {"class": "rendering_person_short"})
for div in content:
    data = div.find_all("a", {"rel": "Person"})
    name_tag = data[0].find_all("span")
    name=name_tag[0].get_text()
    link=data[0]['href']
    d = { 'name':name , 'profile_link':link }
    person.append(d)
current_page=1
while(page_no[1]!= page_no[2]):
    page="https://pureportal.coventry.ac.uk/en/persons?format=&organisationIds=e5cd4a48-01f9-43af-87fb-74738099f357&nofollow=true&page="+str(current_page)
    driver.get(page)  
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    pages =soup.find_all("li", {"class": "search-pager-information"})
    page_no = re.findall('\d+',str(pages[0]))
    content = soup.find_all("div", {"class": "rendering_person_short"})
    for div in content:
        data = div.find_all("a", {"rel": "Person"})
        name_tag = data[0].find_all("span")
        name=name_tag[0].get_text()
        link=data[0]['href']
        d = { 'name':name , 'profile_link':link }
        person.append(d)
    current_page = current_page+1



for p in person:
    driver.get(p['profile_link']+"/publications")  
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    publication =soup.find_all("div", {"class": "result-container"})
    name = p['name']
    profile_link = p['profile_link']
    for div in publication:
        publication_data=div.find_all("h3",{"class":"title"})
        publication_atag=publication_data[0].find_all("a")
        publication_link = publication_atag[0]['href']
        b={ 'name':name , 'profile_link':profile_link , 'publication_link':publication_link}
        books.append(b)
print(len(books))

for bo in books:
    driver.get(bo['publication_link'])  
    name= bo['name']
    profile_link = bo['profile_link']
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    hedding_withtag =soup.find_all("h1")
    abstract_withtag=soup.find_all("div",{"class":"textblock"})
    publication_date_withtag=soup.find_all("span",{"class":"date"})
    key_withtag=soup.find_all("meta",{"name":"citation_keywords"})
    try:
        hedding=hedding_withtag[0].get_text()
    except IndexError:
        hedding = 'null'
    try:
        abstract = abstract_withtag[0].get_text()
    except IndexError:
        abstract = 'null'
    try:
        publication_date = publication_date_withtag[0].get_text()
    except IndexError:
        publication_date = 'null'
    try:
        key1 = key_withtag[0].get_text()
        key = key_withtag[0]['content']
    except IndexError:
        key = 'null'
    data = { 'sl_no':sl_no ,'name':name , 'head_line':hedding ,'keyword':key ,'abstract':abstract ,'profile_link':profile_link,'head_link':bo['publication_link'] ,'published':publication_date}
    sl_no = sl_no+1
    database.append(data)
    print(sl_no)
print("****---------******")
df = pd.DataFrame(database)
df.to_excel("data.xlsx")
print("**** ########## ******")


        
