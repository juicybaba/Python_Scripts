import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

###Test code 
###r = requests.get("http://www.ireadweek.com/index.php/index/"+ "1" +".html")
###c = r.content
###soup = BeautifulSoup(c,"html.parser")
###all = soup.find_all("li")
###len(all)
###all[9:59]


l = []
Currentpage = 1 
MinimumItem = 66 

# Except last page, every page has 66 - 68 itmes.
# Every page has 9 items for heading and 8 item for tail.
# All those items should be eliminated from book list, otherwise, there will be None type error for find().text function.

while(MinimumItem > 65):
    r = requests.get("http://www.ireadweek.com/index.php/index/"+ str(Currentpage) +".html")
    c = r.content
    soup = BeautifulSoup(c,"html.parser")
    all = soup.find_all("li")
    
    if len(all) >= MinimumItem:   # All pages except last page
        print("This is page " + str(Currentpage) + ". This page has " + str(len(all)) + " items.")
        Currentpage += 1
        for item in all[9:59]:
            lst = {}
            lst["Book"] = item.find('div',{'class','hanghang-list-name'}).text
            lst["Download"] = int(item.find('div',{'class','hanghang-list-num'}).text)
            lst["Author"] = item.find('div',{'class','hanghang-list-zuozhe'}).text
            l.append(lst)
            
    else:    #last page
        print("This is page " + str(Currentpage) + "(lastpage). This page has " + str(len(all)) + " items.")
        MinimumItem = len(all)
        for item in all[9:-7]:   # Only add books to the list.
            lst = {}
            lst["Book"] = item.find('div',{'class','hanghang-list-name'}).text
            lst["Download"] = int(item.find('div',{'class','hanghang-list-num'}).text)
            lst["Author"] = item.find('div',{'class','hanghang-list-zuozhe'}).text
            l.append(lst)

df=pd.DataFrame(l)
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()

print("There are "+ str(len(l)-1) + " books in the list.")