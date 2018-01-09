import os
from bs4 import BeautifulSoup
import pandas as pd


class MIDLookup():
    def __init__(self, excel_filename, sheetname, columns, index, year, mid):
        self.excel_filename = excel_filename
        self.sheetname = sheetname
        self.columns = columns
        self.index = index
        self.year = year
        self.mid = mid

    def excel_to_dict(self):
        df = pd.read_excel(
             self.excel_filename,
             sheetname=self.sheetname,
             parse_cols=self.columns,
             converters={'BEGIN YR':str})
        df = df.rename(columns={self.index: 'ISBN', self.year: 'YEAR', self.mid: 'MID'})
        df = df.astype(str)
        df = df.set_index('ISBN')
        lookup = df.to_dict('index')
        return lookup


class XML():
    def __init__(self, xml_filename, mid_lookup):
        self.xml_filename = xml_filename
        self.mid_lookup = mid_lookup
        self.header = '<?xml version="1.0" encoding="UTF-8"?>\n'
        with open(self.xml_filename) as f:
            self.xmlsoup = BeautifulSoup(f, 'lxml')

    def split_xml(self):
        for record in self.xmlsoup.find_all('record'):
            if self.mid_lookup.get(record.datafield.subfield.text) is not None:
                dtformat = self.mid_lookup[record.datafield.subfield.text]['YEAR']
                mid = self.mid_lookup[record.datafield.subfield.text]['MID']
                path = './OUTPUT/' + mid + '/' + dtformat
                fname = record.datafield.subfield.text + '.xml'
                if os.path.exists(path):
                    with open(os.path.join(path, fname), 'w') as f:
                        f.write(self.header + str(record))
                else:
                    os.makedirs(path)
                    with open(os.path.join(path, fname), 'w') as f:
                        f.write(self.header + str(record))

'''
nursQ12017 = MIDLookup('Q1_2017\DissertationsCinahl_Nursing_MARC_Q1_2017_20170519_UpdatePsg.xlsx', 'WorkSheet', 'B,F,L', 'ISBN', 'Start Year', 'MID')
nursQ12017_lookup = nursQ12017.excel_to_dict()
nursQ12017_xml = XML('Q1_2017\Cinahl Nursing MARC Q1 2017.xml', nursQ12017_lookup)
nursQ12017_split = nursQ12017_xml.split_xml()
'''

rehabQ12017 = MIDLookup('Q1_2017\DissertationsCinahl_Rehab_MARC_Q1_2017_20170519_UpdatePsg.xlsx', 'WorkSheet', 'B,D,L', 'ISBN', 'Start Year', 'MID')
rehabQ12017_lookup = rehabQ12017.excel_to_dict()
rehabQ12017_xml = XML('Q1_2017\Cinahl Rehab MARC Q1 2017.xml', rehabQ12017_lookup)
rehabQ12017_split = rehabQ12017_xml.split_xml()
