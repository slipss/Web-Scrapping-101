import urllib2
import json
import requests
import xlwt
import xlrd
import xlutils
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import Workbook
from bs4 import BeautifulSoup
from urllib2 import urlopen



#Function Definitions

# Get temperature of a query q, q needs to be in a string format 'Country/City.json' 
# If needed can get individual weather station reports (For certain well locations)
#Currently in F, can easily change to Celsuis
def get_Current_Temp(q):
    f = urllib2.urlopen('http://api.wunderground.com/api/c30c398d7c3ed667/geolookup/conditions/q/'+q)
    json_string = f.read()
    parsed_json = json.loads(json_string)
    location = parsed_json['location']['city']
    temp_f = parsed_json['current_observation']['temp_f']
    print "Current temperature in %s is: %s" % (location, temp_f)
    f.close()

#Test showing how get_Current_Temp() works
#x='IA/Cedar_Rapids.json'
#get_Current_Temp(x)

#Get the 3 day forecast of a query q, q needs to be in string format 'xxxxxx.json'
def get_Simple_Forecast(q):
    r=requests.get('http://api.wunderground.com/api/c30c398d7c3ed667/forecast/q/'+q)
    data=r.json()

    for day in data['forecast']['simpleforecast']['forecastday']:
        print day['date']['weekday']+':'
        print 'Conditions:',day['conditions']
        print 'High:',day['high']['celsius']+'C','Low:',day['low']['celsius']+'C','\n'
    
#x='Canada/Calgary.json'
#get_Current_Temp(x)
#get_Simple_Forecast(x)

#get and write date to workbook to setup index
def get_Date():
    r=requests.get('http://api.wunderground.com/api/c30c398d7c3ed667/forecast/q/Canada/Calgary.json')
    data=r.json()
    day=data['forecast']['simpleforecast']['forecastday']
    monthName=day[0]['date']['monthname']
    monthNumber=day[0]['date']['month']
    year=day[0]['date']['year']
    dayOfMonth=day[0]['date']['day']
    get_Date.location=2+dayOfMonth
    sheet.write(get_Date.location,0,str(dayOfMonth)+'/'+str(monthNumber)+'/'+str(year),style2)

   
#Get a High and low temp for the day of query q
#order=get_Date.location (for the correct Column)
#Index is the city index Calgary being 1,2 etc.
def get_High_Low(q,order,index):
    r=requests.get('http://api.wunderground.com/api/c30c398d7c3ed667/forecast/q/'+q)
    data=r.json()
    day=data['forecast']['simpleforecast']['forecastday']
    today=day[0]
    High=today['high']['celsius']
    Low=today['low']['celsius']
    sheet.write(order,index, int(High), style1)
    sheet.write(order,index+1, int(Low), style1)
    
#Get the ticker and price from Bloomberg.com URL should be entered 'http://www.bloomberg.com/quote/CL1:COM'
#Note if the price is only 3 digits will take a < in the spread sheet, need to account for this
def get_Price(sectionURL,row):
    html = urlopen(sectionURL).read()
    soup = BeautifulSoup(html, 'lxml')
    priceString = str(soup.find('div','price'))
    price=priceString[priceString.index('>')+1:priceString.index('>')+6]
    if price[-1]=='<':
       price=price[0:4]
    price=float(price)
    sheet.write(get_Date.location,row,price,style1)
    ticker=sectionURL[sectionURL.index('q')+6:]
    sheet.write(2,row,ticker,style1)
    
#Gets the USD value for the Canadian dollar from the noon-day rate
def get_Canadian_Dollar(sectionURL,row):
    html=urlopen(sectionURL).read()
    soup=BeautifulSoup(html,'lxml')
    dollarString=str(soup.find('div','table-responsive'))
    usdString=dollarString[dollarString.index('U.S. dollar'):]
    usdValueLocater=usdString.find('<tr')
    usdValue=usdString[usdValueLocater-16:usdValueLocater-10]
    if usdValue[1]!='.':
       usdValue="N/A"
    else:
        usdValue=float(usdValue)
    sheet.write(get_Date.location,row,usdValue,style1)
    sheet.write(2,row,'USD',style1)
 
#get workbook
rb=open_workbook('december.xls')
wb=copy(rb)
sheet=wb.get_sheet(0)
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')
style1 = xlwt.easyxf()
style2 = xlwt.easyxf(num_format_str='dd-mm-yyyy')

#Website calls
get_Date()
#Calgary Airpot
a='Canada/Calgary.json'
get_High_Low(a,get_Date.location,1)
#Drayton Valley town station
b='zmw:00000.6.71060.json'
get_High_Low(b,get_Date.location,3)
#Goodfare (Hyde town station)
c='locid:CAXX1622;loctype:1.json'
get_High_Low(c,get_Date.location,5)
#Summerland (Peace River town station)
d='zmw:00000.1.71068.json'
get_High_Low(d,get_Date.location,7)
##Lines 1-90 work properly

#Price calls start at row 10, and move on 10-X
e='http://www.bloomberg.com/quote/CRUD:LN'
get_Price(e,10)
f='http://www.bloomberg.com/quote/UNG:US'
get_Price(f,11)
g='http://www.bloomberg.com/quote/SPH:US'
get_Price(g,12)

#Canadian dollar value calls start at 13
h='http://www.bankofcanada.ca/rates/exchange/noon-rates-5-day/'
get_Canadian_Dollar(h,13)

#Saves the document
wb.save('december.xls')
