import requests
import xlwings as xw
import time

api_key = 'e27632d346617c119235b9be5291ede1'
url='http://api.openweathermap.org/data/2.5/weather'

def create_workbook():
    try:
        wb=xw.Book()
    except:
        wb=xw.App().books[0]
    sheet=wb.sheets[0]
    sheet.range('A1:E1').api.Font.Bold=True
    sheet.range('A1:E1').api.Font.Underline=True
    sheet.range('A1:E1').api.Font.Size=18
    sheet.range('A1').value=[['City Name','Temperature','Humidity','UNIT','Update(0/1)']]
    sheet.range('A1:E1').autofit()
    return sheet




def get_weather(city_name):
    param={}
    param['appid']=api_key
    param['q']=city_name
    data=requests.get(url,params=param).json()
    temperature=data['main']['temp']
    humidity=data['main']['humidity']

    return (temperature-273,humidity)



cities=[]

sheet=create_workbook()

def add_new_city(i,city_name):

    temp,humidity=get_weather(city_name)
    sheet.range('A'+str(i+2)).value=[[city_name,temp,humidity,'C',1]]
    time.sleep(0.5)

def update_values():
    check_sheet()
    for i,city in enumerate(sheet.range('A1').expand().value):
        temp=city[1]
        humidity=city[2]
        if city[4]==1.0:
            temp,humidity=get_weather(city[0])
            
        if city[3].lower()=='f':
            tempf=(temp*9/5)+32

            sheet.range('A'+str(i+1)).value=[[city[0],tempf,humidity,city[3],city[4]]]
        else:
            sheet.range('A'+str(i+1)).value=[[city[0],temp,humidity,city[3],city[4]]]
        time.sleep(0.5)



def check_sheet():
    val=sheet.range('A'+str(len(cities)+2)).value
    if val is not None and len(val)>2:
        cities.append(val)
        add_new_city(len(cities)-1,val)

input('Enter City Names in the Excel File and press enter')
print(sheet.range('A2').expand('down').value)
for city in sheet.range('A2').options(expand='down').value:
    cities.append(city)
for i,city in enumerate(cities):
    add_new_city(i,city)
while True:
    update_values()

