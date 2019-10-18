# -*- coding: utf-8 -*-
"""
Created on Thu Sep 26 18:42:44 2019

@author: AHILAN K NATARAJAN
"""


from flask import Flask,send_file
from flask import request
from flask import render_template
#excelsheet
import xlsxwriter
import xlwt
#sorting in excel sheet
import pandas as pd

#for webscrapping
from bs4 import BeautifulSoup as soup
from urllib.request import urlopen as uReq
from urllib.request import urlretrieve
from urllib.parse import quote
from xlwt import Workbook
from xlrd import open_workbook
from collections import Counter

app = Flask(__name__)
@app.route('/')
def my_form():
    return render_template("myform.html")
@app.route('/', methods=['POST'])
def my_form_post():
    text = request.form['text']
    multiply_text = text * 3
    #return multiply_text
    print(text)   
#to create the excel sheet
    workbook = xlsxwriter.Workbook('product.xlsx')
    worksheet = workbook.add_worksheet()
    #headers for the excel sheet
    worksheet.write('A1', 'Keyword')
    worksheet.write('B1', 'URL')
    #filename = "products.csv"
    #to generate search link
    #g=input("Enter your search: ") 
    qstr = quote(text)
    #my_url = "https://www.etsy.com/in-en/search?q=father%20gift"
    my_url = "https://www.etsy.com/in-en/search?q="+qstr
    print(my_url)
    exit=1
    row=1
    while(exit<=1):
        uClient = uReq(my_url)
        page_html1 = uClient.read()
        uClient.close()
        page_soup = soup(page_html1,"html.parser")
        #scrapping all the products present in a page
        containers = page_soup.find("ul",{"class":"responsive-listing-grid wt-grid wt-grid--block justify-content-flex-start pl-xs-0"})     
        print(len(containers))
        c=1
        #scrapping the all the keywords of a product
        for link in containers.findAll("a",{"class":"organic-impression display-inline-block listing-link"}):
            if(c<3):
                my_url1=link.get('href')
                print(my_url1)
                #my_url = "https://www.etsy.com/in-en/listing/713165096/fathers-day-print-custom-photo-gift-idea?ga_order=most_relevant&ga_search_type=all&ga_view_type=gallery&ga_search_query=fathers+day+gift&ref=sr_gallery-1-1&organic_search_click=1&frs=1"
                uClient = uReq(my_url1)
                page_html = uClient.read()
                uClient.close()
                page_soup = soup(page_html,"html.parser")
                containers = page_soup.findAll("li", {"class":"list-inline-item wt-mb-xs-1 wt-mr-xs-1"})
                #row=1
                col=0
                for container in containers:
                    keyw=container.text
                    print(container.text)
                    link_finder= container.find("a",{"class":"btn btn-secondary"})
                    url=link_finder.get('href')
                    print(keyw + "," +url)
                    worksheet.write(row,0,keyw)
                    worksheet.write(row,1,url)
                    row += 1
                c+=1  
            else:
                 break
        #uClient = uReq(my_url)
        #page_html = uClient.read()
        #uClient.close()
        page_soup = soup(page_html1,"html.parser")
        next_page=page_soup.find("ul",{"class":"wt-action-group wt-list-inline"})
        link_finder=next_page.find("a",{"class":"wt-btn wt-btn--small wt-action-group__item wt-btn--icon"})

        print(link_finder)
        #to end if the search result does not contain 2nd page
        if link_finder:
            my_url=url=link_finder.get('href')
        else:
            exit=3
        exit+=1

    workbook.close()

    

    #to find the count of each keyword

    book=open_workbook("product.xlsx")
    sheet=book.sheet_by_index(0)
    l=[]
    wb=Workbook()
    sheet1=wb.add_sheet('Sheet1')
    for k in range(1,sheet.nrows):
        l.append(str(sheet.row_values(k)[0]))
    print(l)
    counts=Counter(l)
    print(counts)
    i=1
    j=0
    #store the keywords with count in excel sheet
    sheet1.write(0,0,"Keyword")
    sheet1.write(0,1,"Count")
    for key in counts:
        print(key,counts[key])
        sheet1.write(i,0,key)
        sheet1.write(i,1,counts[key])
        i=i+1
    wb.save('result.xlsx')
    xl=pd.ExcelFile("result.xlsx")
    df=xl.parse("Sheet1")

    #to sort the keywords based on count
    df=df.sort_values(by='Count',ascending=False)
    writer=pd.ExcelWriter('result.xlsx')
    df.to_excel(writer,sheet_name='Sheet1',columns=["Keyword","Count"],index=False)
    writer.save()
    #book.close()
    print(type(counts))
    try:
     	return render_template('downloads.html')
    except Exception as e:
    	return str(e)
    
@app.route('/return-files/')
def return_files_tut():
	try:
		return send_file('result.xlsx')
	except Exception as e:
		return str(e)

if __name__ == '__main__':
    app.run()
