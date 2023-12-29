import bs4, pandas as pd, openpyxl, requests, re
from openpyxl.utils.dataframe import dataframe_to_rows

results = requests.get(r'https://www.stateauto.com/findanagent?lo=-77.67651839999999&la=38.1906846&insuranceType=Businesses&zip=22551')
print(results.status_code == requests.codes.ok)

soupy = bs4.BeautifulSoup(results.text, 'html.parser')
agency_df = pd.DataFrame({})

def html_parser(soup_object,col_name):
    object = soupy.select(soup_object)
    obj_list = []
    for i in object:
        obj_list.append(i.get_text())
    agency_df[col_name] = obj_list

agencies = soupy.select('strong')

html_parser('strong','Agencies')

#business type
business_type = soupy.find_all('div',class_='col-lg-2 col-md-2 col-sm-2 col-xs-11')
business_type = [str(i) for i in business_type]
business_form = re.compile(r'<p>(.*)</p>')
business = [business_form.search(x).group(1) if business_form.search(x) else '' for x in business_type]
agency_df['Business Type'] = business

#info block has phone numbe and address
info = soupy.find_all('div',class_='col-lg-3 col-md-3 col-sm-3 col-xs-11')
print(len(info))
info = [str(i) for i in info]

#extracting phone using regex:
number_form = re.compile(r'Phone: (\(\d{3}\) \d{3}-\d{4})')
phone_numbers = [number_form.search(x).group(1) if number_form.search(x) else '' for x in info]
agency_df['Phone Numbers'] = phone_numbers

#extracting adress using regex:
address_form = re.compile(r'(.*)Phone')
address = [address_form.search(x).group(1) if address_form.search(x) else '' for x in info]
address = [a.strip().replace('<br/>','\n') for a in address]
agency_df['Address'] = address 

print(agency_df)

#print to excel w/ openpyxl
wb = openpyxl.load_workbook(r'misc.-files\Insurance_scrape.xlsx')
sheet = wb['Sheet']
sheet.title = 'State Auto'

rows = dataframe_to_rows(agency_df,index=False, header=True)

for r_idx, row in enumerate(rows, 0):
    for c_idx, value in enumerate(row, 0):
         sheet.cell(row=r_idx+1, column=c_idx+1, value=value)

wb.save('misc.-files\Insurance_scrape.xlsx')