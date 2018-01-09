from bs4 import BeautifulSoup
import xlsxwriter
import pandas as pd

filename = 'Q1_2017\Cinahl Rehab MARC Q1 2017.xml'
metadata_file = open(filename, encoding='utf8').read()
xmlsoup = BeautifulSoup(metadata_file, 'lxml')

xml_dict = {}

for record in xmlsoup.find_all('record'):
    for an_tag in record.find('controlfield', {'tag': '001'}):
        xml_dict[an_tag] = {}
    for isbn_tag in record.find_all('datafield', {'tag': '020'}):
        xml_dict[an_tag]['ISBN'] = isbn_tag.find('subfield', {'code': 'a'}).text
    for title_tag in record.find_all('datafield', {'tag': '245'}):
        xml_dict[an_tag]['Title'] = title_tag.find('subfield', {'code': 'a'}).text  
    for year_tag in record.find_all('datafield', {'tag': '502'}):
        xml_dict[an_tag]['Year'] = year_tag.find('subfield', {'code': 'd'}).text

df = pd.DataFrame.from_dict(xml_dict, orient='index')
excel_writer = pd.ExcelWriter('Cinahl_Rehab_MARC_Q1_2017.xlsx', engine='xlsxwriter')
df.to_excel(excel_writer, index_label='AN')
excel_writer.save()