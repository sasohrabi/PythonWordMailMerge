from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

import os

import sqlite3
import pandas as pd

con = sqlite3.connect('C:/output/checkrecived.sqlite')
c = con.cursor()
c.execute("select * from tblPartyAccount")
data = pd.read_sql_query("select code,PartyAccount from tblPartyAccount", con)
##print(data['code'].values)

for row in data.itertuples():
    print (row)

template = "c:/output/word/fac.docx"
document = MailMerge(template)
print(document.get_merge_fields())
##document.merge(namef='Gold')
fa = [{'nam': 'sa', 'fam': 'so'},
{'nam': 'se', 'fam': 'so'},
{'nam': 'fa', 'fam': 'za'}]
document.merge_rows('nam', fa)
document.write('c:/output/word/doc4.docx')