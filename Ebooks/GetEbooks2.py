import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

l = []
Currentpage = 1 
MinimumItem = 51

###Test code    1111 
#r = requests.get("http://www.ireadweek.com/index.php/index/"+ "196" +".html")
#c = r.content
#soup = BeautifulSoup(c,"html.parser")
#name = soup.find_all("div",{"class":"hanghang-list-name"})
#zuozhe = soup.find_all("div",{"class":"hanghang-list-zuozhe"})
#num = soup.find_all("div",{"class":"hanghang-list-num"})
#len(name)

while(MinimumItem == 51):
    r = requests.get("http://www.ireadweek.com/index.php/index/"+ str(Currentpage) +".html")
    c = r.content
    soup = BeautifulSoup(c,"html.parser")
    name = soup.find_all("div",{"class":"hanghang-list-name"})
    zuozhe = soup.find_all("div",{"class":"hanghang-list-zuozhe"})
    num = soup.find_all("div",{"class":"hanghang-list-num"})
    print("This is page " + str(Currentpage) + ". This page has " + str(len(name)) + " items.")
    
    for nm,zz,nu in zip(name,zuozhe,num):
        lst = {}
        lst["Book"] = nm.text
        lst["Download"] = nu.text
        lst["Author"] = zz.text
        l.append(lst)
    Currentpage += 1

df=pd.DataFrame(l)
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()

print("There are "+ str(len(l)-1) + " books in the list.")