'''
Module that uses JSON data from http://country.io/data/ and builds an CSV File
This code is for educational purpose
Working of program depends on the data availability in above mentioned URL
Related Kaggle URL for same data - https://www.kaggle.com/timoboz/country-data
'''
import json #To Deal with Json Data
import urllib.request #To read Data from URL
import csv #To combine Whole 5 JSON Data files to Excel
import datetime #To Generate Unique File names

'''Data Source URL's'''
CONTINENT_URL = "http://country.io/continent.json" #Based on ISO2 Codes
COUNTRY_NAMES_URL = "http://country.io/names.json" #Based on ISO2 Codes
COUNTRY_ISO3_URL = "http://country.io/iso3.json" #ISO2 to ISO3 Mapping
COUNTRY_CAPITAL_URL = "http://country.io/capital.json" #Based on ISO2 Codes
COUNTRY_PHONE_URL = "http://country.io/phone.json" #Based on ISO2 Codes
COUNTRY_CURRENCY_URL = "http://country.io/currency.json" #Based on ISO2 Codes
'''End of Data Source URL's'''

'''Read Data from URL and Store them in Variables'''
continent_content = urllib.request.urlopen(CONTINENT_URL)
CONTINENT_DATA = json.loads(continent_content.read())
country_names_content = urllib.request.urlopen(COUNTRY_NAMES_URL)
COUNTRY_NAMES_DATA = json.loads(country_names_content.read())
iso2_iso3_content = urllib.request.urlopen(COUNTRY_ISO3_URL)
COUNTRY_ISO2_ISO3 = json.loads(iso2_iso3_content.read())
country_capital_content = urllib.request.urlopen(COUNTRY_CAPITAL_URL)
CAPITAL_DATA = json.loads(country_capital_content.read())
country_phone_content = urllib.request.urlopen(COUNTRY_PHONE_URL)
PHONE_DATA = json.loads(country_phone_content.read())
country_currency_content = urllib.request.urlopen(COUNTRY_CURRENCY_URL)
CURRENCY_DATA = json.loads(country_currency_content.read())
'''End of Read Data from URL and Store them in Variables'''

'''Get the countries codes in Alphabetical Order'''
countries =  list(COUNTRY_NAMES_DATA.keys())
countries.sort() #Sort the Country Code Alphabetically
'''End of Get the countries codes in Alphabetical Order'''

#Header for Excel File Output
HEADER = [
    'COUNTRY NAME',
    'CONTINENT NAME',
    'COUNTRY ISO2 CODE',
    'COUNTRY ISO3 CODE',
    'CAPITAL',
    'CURRENCY CODE',
    'PHONE CODE',
    ]

#TimeStamp for Unique file generation
timestamp = datetime.datetime.now().strftime("%d%m%y%H%M%S")

'''CSV File related Objects'''
csv_file = open("COUNTRY_DETAILS" + timestamp + ".csv",'w',newline='')
csv_writer = csv.DictWriter(csv_file,delimiter=',',fieldnames=HEADER)
'''End of CSV File Related Objects'''

#Write Header for CSV File
csv_writer.writeheader()

#Write Rows in CSV File
for key in countries:
    values = [
              COUNTRY_NAMES_DATA[key],
              CONTINENT_DATA[key],
              key,
              COUNTRY_ISO2_ISO3[key],
              CAPITAL_DATA[key],
              CURRENCY_DATA[key],
              PHONE_DATA[key],
              ]
    row = dict(zip(HEADER,values))
    csv_writer.writerow(row)
csv_file.close()
