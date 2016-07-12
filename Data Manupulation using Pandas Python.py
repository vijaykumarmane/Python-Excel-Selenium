""" Data manupulation like excel and Importing Exporting data using Pandas.
And function like IF and VLookup*, adding new column and filling using lambda function like formulaes in excel and
deleting and updating data, creating new dataframes like new sheet."""

import pandas as pd
import numpy as np

# Load Data into Python memory.
path = file_path
with pd.ExcelFile(path) as xls:
    df = pd.read_excel(xls, Sheet_name )
# Strip headers of all columns
df = df.rename(columns=lambda x: x.strip())

# Code to remove spaces in header with underscore. You can use like
# df2.columns = [x.strip().replace(' ', '_') for x in df2.columns]

# Cleaning for PT Count. "DN / PT#" is a column name in dataset.
df['DN / PT#'] = df['DN / PT#'].replace(', ',',')
df['DN / PT#'] = df['DN / PT#'].replace('No D/N on docs','')

# Enter PT Count
df['PT COUNT'] = df['DN / PT#'].apply(lambda x: str(x).count(',')+1 if str(x).count(',') != 0 else (1 if len(str(x))!=0 else 0 ))

# Pick up Delay Yes or No
def myTest(row):
    try:
        result = row['Pickup Cutoff Time'] > row['Actual Pickup Time']
    except:
        result = False
    if result:
        return('N')
    else:
        return('Y')
df['Late Pickup Y/N?'] = df[['Pickup Cutoff Time','Actual Pickup Time']].apply(lambda row: myTest(row), axis=1)

# Delivery Delay Stat
df['Reason Type'] = df['Reason Type'].str.title()
df['Delivery Delay Days'] = df['EDD'] - df['POD Date']

def test(x):
    if x > pd.Timedelta(0):
        x = 'Late'
    elif x == pd.Timedelta(0):
        x = 'On Time'
    elif x < pd.Timedelta(0):
        x = 'Early'
    else:
        x = 'No POD'
    return x


df['Delivery Time Status'] = df['Delivery Delay Days'].apply(lambda x : test(x));


# Comparing with Reason type.
def test2(row1,row2):
    if row1 == 'Late' and row2 == 'Uncontrollable':
        row = 'On Time'
    else:
        row = row1
    return row

df['Delivery Time Status'] = df[['Delivery Time Status','Reason Type']].apply(lambda row : test2(
        row['Delivery Time Status'],row['Reason Type']),axis = 1)
        
# Cleaning Data Dates
df['ATD'] = df['ATD'].replace('-',pd.NaT);
df['ETD'] = df['ETD'].replace('-',pd.NaT);

df['ATD-ETD'] = df['ETD'] - df['ATD'];
df['ATD-ETD'] = df['ATD-ETD'].fillna(pd.NaT)
def test(x):
    if x > pd.Timedelta(0):
        x = 'Late'
    elif x == pd.Timedelta(0):
        x = 'On Time'
    elif x < pd.Timedelta(0):
        x = 'Early'
    else:
        x = 'No Info'
    return x
df['DEPARTURE STATUS'] = df['ATD-ETD'].apply(lambda x : test(x));

# For Week Numbering for Quarter starting from 1
df["Week"]= df['Pickup date'].dt.week.apply(lambda x: x-12)

# Correcting row values to Title format
df['Mode'] = df['Mode'].str.title()

path = vlook_up_file_path
with pd.ExcelFile(path) as xls1:
    rdf = pd.read_excel(xls1, 'Sheet1')
# Strip headers
rdf = rdf.rename(columns=lambda x: x.strip())

# Vlook up
df = df.merge(rdf[['ISO (2)','Continent']], how='left',
         left_on='Destination country', right_on='ISO (2)',
         left_index=False, right_index=False)

# Removal of Extra column and 
df = df.drop('ISO (2)', axis=1)
df['Continent'] = df['Continent'].fillna('Other')

# Creating OTP raw data.
week_arr = df['Week'].unique();week_arr.sort()
mode_arr = df['Mode'].unique(); mode_arr.sort()
regiArr = df['Continent'].unique();regiArr.sort()
ntdf = pd.DataFrame(columns=('Region','Week','Mode'))
dft =[]
for regi in regiArr:
    temp = [];
    for week in week_arr:
        for mode in mode_arr:
            temp = [regi,week, mode]
            dft.append(temp)
otp = pd.DataFrame(dft,columns=['Continent','Week','Mode'])

# Temp DataFrame
tdf = df[['Continent','Week','Mode','Reason Type','Type','Late Pickup Y/N?','Delivery Time Status']]
# Function to Calculate Week OTP Gross and Net
def otpCal(region,mode,week):
    dEarly = tdf['Type'][tdf['Delivery Time Status'] == 'Early'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max();
    dLate = tdf['Type'][tdf['Delivery Time Status'] == 'Late'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max();
    dOntime = tdf['Type'][tdf['Delivery Time Status'] == 'On Time'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max();
    dNopod = tdf['Type'][tdf['Delivery Time Status'] == 'No POD'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max();
    ftf = tdf['Mode'][tdf['Type'] == 'FTF'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max()
    Imports = tdf['Type'][tdf['Type'] == 'Imports'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max()
    Exports = tdf['Type'][tdf['Type'] == 'Exports'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max()
    Domestic = tdf['Type'][tdf['Type'] == 'Domestic'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max()
    Transborder = tdf['Type'][tdf['Type'] == 'Transborder'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max()
    Cnt = tdf['Week'][tdf['Delivery Time Status'] != 'No POD'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode][tdf['Reason Type'] == 'Controllable'].count().max()
    Uncnt = tdf['Week'][tdf['Delivery Time Status'] != 'No POD'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode][tdf['Reason Type'] == 'Uncontrollable'].count().max()
    wk1tt = tdf['Week'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max()
    lateY = tdf['Type'][tdf['Late Pickup Y/N?'] == 'Y'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max();
    lateN = tdf['Type'][tdf['Late Pickup Y/N?'] == 'N'][tdf['Continent'] == region][tdf['Week'] == week][tdf['Mode'] == mode].count().max();
    netOtp= ((wk1tt-Cnt-dNopod)/(wk1tt-dNopod))*100
    grossOtp = ((wk1tt-(Cnt+Uncnt+dNopod))/(wk1tt-dNopod))*100
    return(netOtp,grossOtp,wk1tt,Cnt,Uncnt,Imports,Exports,Domestic,ftf,Transborder,lateY,lateN,dEarly,dLate,dOntime,dNopod)
    
# Calling function
otpArr = otp.apply(lambda row: otpCal(row['Region'],row['Mode'],row['Week']), axis=1)

# Adding columns to OTP DataFrame
otp['Total Shipment'] = otpArr.apply(lambda x: x[2]);
otp['%Net_OTP'] = otpArr.apply(lambda x: x[0]); 
otp['%Gross_OTP'] = otpArr.apply(lambda x: x[1]);
otp['Controllable'] = otpArr.apply(lambda x: x[3]);
otp['Uncontrollable'] = otpArr.apply(lambda x: x[4]);
otp['Imports'] = otpArr.apply(lambda x: x[5]);
otp['Exports'] = otpArr.apply(lambda x: x[6]);
otp['Domestic'] = otpArr.apply(lambda x: x[7]);
otp['FTF'] = otpArr.apply(lambda x: x[8]);
otp['Transborder'] = otpArr.apply(lambda x: x[9]);
otp['Late Pickup'] = otpArr.apply(lambda x: x[10]);
otp['On Time Pickup'] = otpArr.apply(lambda x: x[11]);
otp['Early Delivery'] = otpArr.apply(lambda x: x[12]);
otp['Late Delivery'] = otpArr.apply(lambda x: x[13]);
otp['On Time Delivery'] = otpArr.apply(lambda x: x[14]);
otp['No POD'] = otpArr.apply(lambda x: x[15]);

# OTP DataFrame to Excel using Openpyxl liabrary
from openpyxl import load_workbook
otp = otp[otp['Total Shipment'] != 0]
otp = otp.copy()
book = load_workbook(file_path)
writer = pd.ExcelWriter(file_path,
                        engine='openpyxl') 
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
otp.to_excel(writer,sheet_name="Data",na_rep=0,index=False)
writer.save()
