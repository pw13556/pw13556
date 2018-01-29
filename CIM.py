# coding: utf-8
from datetime import datetime, timedelta
import pandas as pd
import csv
import openpyxl

# import and clean
df = pd.read_excel(open('CIM.xlsx', 'rb'), sheet_name='Report', index_col=14)  # index_col=26
df['Source'].replace(["Group Internal Audit Department (GIAD) audit findings",
                      "Management Self-Identified Control Issues (MSII), including routine RCSA & ShARP testing",
                      "Loss events' root cause analysis", 'Regulatory audit findings', 'Others',
                      'Compliance Monitoring Review Findings'],
                     ['Audit', 'MSII', 'LED', 'Others', 'Others', 'Others'], inplace=True)

df['Organisation Level'].replace(["CIMB THAI>Group Information and Operations Division>Technology Division>",
                                  "CIMB THAI>Group Information and Operations Division>Technology Division>IT Security Management and Governance Department>",
                                  "CIMB THAI>Group Information and Operations Division>Technology Division>IT Security Management and Governance Department>IT Security Department>",
                                  "CIMB THAI>Group Information and Operations Division>Technology Division>IT Infrastructure Management Department>",
                                  "CIMB THAI>Group Information and Operations Division>Operations Division>",
                                  "CIMB THAI>Group Information and Operations Division>Operations Division>Credit Operation Department>",
                                  "CIMB THAI>Group Information and Operations Division>Operations Division>Credit Administration Team>",
                                  "CIMB THAI>Group Information and Operations Division>Technology Division>Programme Delivery Department>",
                                  "CIMB THAI>Group Information and Operations Division>Operations Division>Capital Financial Markets & Payments Operation Department>",
                                  "CIMB THAI>Group Information and Operations Division>Technology Division>Application Management Department>",
                                  "CIMB THAI>Group Information and Operations Division>Operations Division>Domestic Banking Department>",
                                  "CIMB THAI>Group Information and Operations Division>Operations Division>International Business Department>",
                                  "CIMB THAI>Group Information and Operations Division>GIOD Planning and Risk Management Team>Dome and Process Quality and Control Team>",
                                  "CIMB THAI>Group Information and Operations Division>Technology Division>Enterprise Architecture and System Integrtion Department >"],
                                 ["IT", "ITSG", "ITSG", "ITIM", "OP", "CO", "CO", "PGD", "CPO", "APM", "DB", "INB",
                                  "I&O RA", "EAS"], inplace=True)
df['Classification'].replace(['High', 'Medium', 'Low'], ['1High', '2Medium', '3Low'], inplace=True)

# Map dept
Deptname = {"IT": 'IT', "ITSG": 'IT', "ITIM": 'IT', "OP": 'Ops', "CO": 'Ops', "CAT": 'Ops', "PGD": 'IT',
            "CPO": 'Ops', "APM": 'IT', "DB": 'Ops', "INB": 'Ops', "I&O RA": 'I&O RA', "EAS": 'IT'}
df['Country'] = df['Organisation Level'].map(Deptname)

# rename column header
df = df.rename(
    columns={'Country': 'Division', 'SNC': 'PassD', 'Islamic Business': 'Pass Due?', 'Tag': 'y.notificationdate',
             'Tag1': 'y.issueindentifiedate',
             'Tag2': 'Tag.12mbackward', 'Tag3': 'm.today', 'Tag4': 'datetoday', 'Tag5': 'm.diff'})

# Insert column
df.insert(3, 'Audit Issue No.', 1)
df.insert(4, 'Sumdis', 0)

# Map Audit ID
reader = csv.reader(open('MapIDAudit.csv', 'r'))
d = {}
for row in reader:
    k, v = row
    d[k] = v

df['Audit Issue No.'] = df['Issue ID'].map(d)

# calc pass due
today = datetime.today()
df['PassD'] = df['Issue Resolution Date'] < today
Passdue = {True: 'Over due', False: 'Not due'}
df['Pass Due?'] = df['PassD'].map(Passdue)

# calc date 12 month backward ---------------------------------------------------xxx
df['datetoday'] = today
N = 365
date_N_days_ago = today - timedelta(days=N)

# get year/month form date object
df['y.notificationdate'] = pd.DatetimeIndex(df['Notification Date']).year
df['y.issueindentifiedate'] = pd.DatetimeIndex(df['Date Issue Identified']).year

# pivoting
df1 = pd.pivot_table(df, values='Issue ID', index='y.notificationdate', columns='Source',
                     aggfunc=pd.Series.nunique, )  # pd.Series.nunique
df2 = pd.pivot_table(df, values='Issue ID', index='y.notificationdate', columns=['Pass Due?', 'Classification'],
                     aggfunc=pd.Series.nunique, )  # pd.Series.nunique
df0 = pd.pivot_table(df, values='Issue ID', index=['Source', 'Issue Status', 'Pass Due?'], columns='Classification',
                     aggfunc=pd.Series.nunique, )  # pd.Series.nunique
df3 = df0.xs(('MSII', 'Open'), level=0)
df4 = pd.pivot_table(df, values='Issue ID', index='y.notificationdate', columns=['Pass Due?', 'Classification'],
                     aggfunc=pd.Series.nunique, )  # pd.Series.nunique

# export to excel
writer = pd.ExcelWriter('PivotCIM.xlsx')
df.to_excel(writer, 'raw-data')
df1.to_excel(writer, 'Torch P13.1')
df2.to_excel(writer, 'Torch P13.2')
df3.to_excel(writer, 'OpsRisk - MSII Open')
writer.save()

# insert sumdis formula
wb = openpyxl.load_workbook('PivotCIM.xlsx')
sheet = wb['raw-data']
sheet['F2'] = '=IF(COUNTIF(B2:$B$12410,B2)>1,0,1)'
wb.save('PivotCIM.xlsx')
