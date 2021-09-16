import pandas as pd
import numpy as np
from datetime import time
from datetime import date
import datetime
import os
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from timeit import default_timer as timer

global InitialDir
InitialDir = 'C:/'
# ============^^^^^^^^^==================IMPORT BIBLIOTEK============^^^^^^^^^=================


# ============vvvvvvvvvv=================OPENING FILES AND GETTING PATHS============vvvvvvvvvv=================

def GetListFilePath():
    global InitialDir  
    root.filename = filedialog.askopenfilename(initialdir=InitialDir,title='Choose lists file', filetypes=[('Excel files','.xlsx .xlsm .xlsb')])
    if root.filename != '':
        listsEntry.delete(0,END)
        listsEntry.insert(0,root.filename)
    InitialDir = root.filename
    InitialDir = InitialDir[:InitialDir.rfind('/')+1]

def GetRawDataFilePath():
    global InitialDir  
    root.filename = filedialog.askopenfilename(initialdir=InitialDir,title='Choose raw data file', filetypes=[('Excel files','.xlsx .xlsm .xlsb')])
    if root.filename != '':
        rawDataEntry.delete(0,END)
        rawDataEntry.insert(0,root.filename)
    InitialDir = root.filename
    InitialDir = InitialDir[:InitialDir.rfind('/')+1]

def GetNewFilePath():
    global InitialDir  
    root.filename = filedialog.askdirectory(initialdir=InitialDir,title='Choose destination folder directory',)
    if root.filename != '':
        newPathEntry.delete(0,END)
        newPathEntry.insert(0,root.filename + '/')
    InitialDir = root.filename

def GetMappingFilePath():
    global InitialDir  
    root.filename = filedialog.askopenfilename(initialdir=InitialDir,title='Choose your Mapping file', filetypes=[('Excel files','.xlsx .xlsm .xlsb')])
    if root.filename != '':
        mappingEntry.delete(0,END)
        mappingEntry.insert(0,root.filename)
    InitialDir = root.filename
    InitialDir = InitialDir[:InitialDir.rfind('/')+1]

# ============^^^^^^^^^==================OPENING FILES AND GETTING PATHS============^^^^^^^^^=================

# ============vvvvvvvvvv=================UŻYWANE CLASSY / FUNKCJE============vvvvvvvvvv=================

class loadXlsxTable():
    def __init__(self,filePath):
        #Wczytywanie tabeli
        self.table = pd.read_excel(filePath)
        self.table = self.table.iloc[:50,2:]
        #Wyznaczanie indexu
        idx = np.flatnonzero(self.table[self.table.columns[0]].notnull())
        #Wczytywanie tabeli ponownie, z poprawnymi nagłówkami
        self.table = pd.read_excel(filePath,header=idx[0]+1)
        #Wytnij 2 pierwsze kolumny
        self.table = self.table.iloc[:,2:]
        
    def describeTable(self):
        print(self.table.info())
        
    def getDF(self):
        return self.table
        
    def cleanNoHeaders(self):
        print('List of columns befor:')
        self.columnsList = self.prepareColumnsList()
        self.newColumnsList = []
        
        for col in self.columnsList:
            if 'Unnamed:' not in col:
                self.newColumnsList.append(col)
        self.table = self.table[self.newColumnsList]
        print('List of columns after:')
        print(self.newColumnsList)
    
    def prepareColumnsList(self):
        print(list(self.table.columns))
        return list(self.table.columns)
    
    def describeColumns(self,columnNames):
        print(self.table[columnNames].describe())
        
    def findHeaders(self,i=15,j=5):
        print(self.table.iloc[0:i,0:j])


def countBusinessDays(start, end,holidays=[]):
    mask = pd.notnull(start) & pd.notnull(end)
    
    if start.dtype == 'datetime64[ns]':
        start = start.values.astype('datetime64[D]')[mask]
    elif start.dtype == 'object':
        start = start.astype(str).str.split(' ',n = 1,expand=True )[mask]
        start = start[0].values.astype('datetime64[D]')

    if end.dtype == 'datetime64[ns]':
        end = end.values.astype('datetime64[D]')[mask]
    elif end.dtype == 'object':
        end = end.astype(str).str.split(' ',n = 1,expand=True )[mask]
        end = end[0].values.astype('datetime64[D]')
        
    result = np.empty(len(mask), dtype=float)
    result[mask] = np.busday_count(start, end,holidays=holidays)
    result[~mask] = np.nan
    return result


def nextBusinessDay(startDay, businessDays,holidays):
    """This function returns the next businness day
    in datetime format. It based on date from given column 'startDay' value 
    and days to add in another column 'businessDays'.
    """
    if len(holidays) == 0:
        holidays = ''
    def calculateFinalDay(startDay, businessDays,holidays):
        ONE_DAY = np.timedelta64(1,'D')
        final_days = np.empty(0,dtype='datetime64[ns]')
        for i,j in zip(startDay,businessDays):
            temp_day = i
            for x in range(0, int(j)):
                temp_day = pd.to_datetime(temp_day)
                next_day = temp_day + ONE_DAY
                while next_day.weekday() in [5,6] or next_day in holidays:
                    next_day += ONE_DAY
                temp_day = next_day
            temp_day = np.datetime64(temp_day)
            final_days = np.append(final_days,temp_day)
        return final_days

    mask = pd.notnull(startDay) & pd.notnull(businessDays)
    
    if startDay.dtype == 'datetime64[ns]':
        startDay = startDay.values.astype('datetime64[ns]')[mask]
    elif startDay.dtype == 'object':
        startDay = pd.to_datetime(startDay[mask])
    
    if businessDays.dtype == 'str':
        businessDays = businessDays.values.astype('int')[mask]
    elif businessDays.dtype == 'object':
        businessDays = businessDays.values.astype('int')[mask]
    else:
        businessDays = businessDays[mask]

    result = np.empty(len(mask), dtype='datetime64[ns]')
    result[~mask] = np.datetime64("NaT")
    result[mask] = calculateFinalDay(startDay, businessDays,holidays)
    return result


def create_folder(dir):
    dir_path = f'{dir}{date.today()}'
    x = 0 
    while True:
        if not os.path.exists(dir_path):
            os.mkdir(dir_path)
            dir_path = f'{dir_path}/'
            break
        elif not os.path.exists(f'{dir_path}_{x}'):
            os.mkdir(f'{dir_path}_{x}')
            dir_path = f'{dir_path}_{x}/'
            break
        else:
            x += 1
    return dir_path

# ============^^^^^^^^^^============UŻYWANE CLASSY / FUNKCJE============^^^^^^^^^^============

#============vvvvvvvvvv============LOADING FILES============vvvvvvvvvv============

def loadingFiles(rawDataPath,listsFilePath,mappingFilePath):

#Assigning paths
    mappingFilePath = mappingFilePath
    listsFilePath =listsFilePath
    rawDataPath =rawDataPath

#Raw Data loading
    rawDataFile = loadXlsxTable(rawDataPath)
    rawDataFile.cleanNoHeaders()
    rawDataFile = rawDataFile.getDF()

#Lists File Loading
    listsFilePreFix = pd.read_excel(listsFilePath,sheet_name='Pre-Fix')
    listsFileAirportOrigin = pd.read_excel(listsFilePath,sheet_name='AirportCity_Origin')
    listsFileSL = pd.read_excel(listsFilePath,sheet_name='SL')
    listsFileCountry = pd.read_excel(listsFilePath,sheet_name='Country')
    listsFileRegion = pd.read_excel(listsFilePath,sheet_name='Region')
    listsFileMain =  pd.read_excel(listsFilePath,sheet_name='Main')
    listsFileHolidays = pd.read_excel(listsFilePath,sheet_name='Holidays')
    listsFileDelayCodes = pd.read_excel(listsFilePath,sheet_name='Delay_codes')

# Mapping file Loading
    mapFile = pd.read_excel(mappingFilePath)
    # Stworzenie listy kolumn z pliku z QV
    rawDataColumns = mapFile['Raw_names'].tolist()
    rawDataColumns = [x.strip() for x in rawDataColumns]

    # Stworzenie listy kolumn o docelowej nazwie      
    finalDataColumns = mapFile['Final_names'].tolist()
    finalDataColumns = [x.strip() for x in finalDataColumns]

#Creating dictionaries
    listsFilePreFix['Pre-Fix'] = listsFilePreFix['Pre-Fix'].astype(str)
    listsFilePreFix['Pre-Fix'] = [f'00{x}' if len(x) == 1 else f'0{x}' if len(x) == 2 else x for x in listsFilePreFix['Pre-Fix']]
    prefixDict = pd.Series(listsFilePreFix['code'].values,index=listsFilePreFix['Pre-Fix']).to_dict()

    countryRegionDict = pd.Series(listsFileMain['Region'].values,index=listsFileMain['Country']).to_dict()

    SLDict = pd.Series(listsFileSL['code'].values,index=listsFileSL['SL']).to_dict()

    gmtDict = pd.Series(listsFileMain['Local -> GMT'].values,index=listsFileMain['Airport']).to_dict()

    delayDescriptionDict = pd.Series(listsFileDelayCodes['Description'].values,index=listsFileDelayCodes['Code']).to_dict()

    delayControllableDict = pd.Series(listsFileDelayCodes['Controllable'].values,index=listsFileDelayCodes['Code']).to_dict()

#Return
    return rawDataFile, prefixDict, countryRegionDict, SLDict,gmtDict,listsFileHolidays,delayDescriptionDict,delayControllableDict,rawDataColumns,finalDataColumns


# ============^^^^^^^^^^============LOADING FILES============^^^^^^^^^^============



def prepareRawDataASML(rawDataFile,prefixDict,countryRegionDict,SLDict,gmtDict,listsFileHolidays,delayDescriptionDict,delayControllableDict,rawDataColumns,finalDataColumns):


#Clean for only ASML DATA
    rawDataFile = rawDataFile[rawDataFile[['Consignor Name','Consignee Name']].stack().str.contains('ASML|HQ|Hermes|HMI').any(level=0)]

#Airline
    rawDataFile['Airline'] = [prefixDict.get(str(x)[:3],'') if x != np.nan else '' for x in rawDataFile['Master']]

# Actual Pickup
    rawDataFile['Actual Pickup'] = np.where(rawDataFile['Actual Pickup'].isna(),
                                        rawDataFile['Event (IRP - %)'],
                                        rawDataFile['Actual Pickup'])

#Orig. Region
    rawDataFile['Orig. Region'] = [countryRegionDict.get(x,'') if x != np.nan else '' for x in rawDataFile['Origin Ctry']]

#Orig. Airport
    rawDataFile['Orig. Airport'] = rawDataFile['Origin'].str[-3:]

#Dest. Region
    rawDataFile['Dest. Region'] = [countryRegionDict.get(x,'') if x != np.nan else '' for x in rawDataFile['Dest Ctry']]

#Dest. Airport
    rawDataFile['Dest. Airport'] = rawDataFile['Destination'].str[-3:]

#Delivery Terms
    conditions = [
    (rawDataFile['Cartage Pickup Mode'] =='DSV') & (rawDataFile['Cartage Drop Mode'] == 'DSV'), #DD
    (rawDataFile['Cartage Pickup Mode'] =='DSV') & (rawDataFile['Cartage Drop Mode'] == 'CCX'), #DP
    (rawDataFile['Cartage Pickup Mode'] =='DLV') & (rawDataFile['Cartage Drop Mode'] == 'DSV'), #PD
    (rawDataFile['Cartage Pickup Mode'] =='DLV') & (rawDataFile['Cartage Drop Mode'] == 'CCX')  #PP
             ]
    choices = ['DD','DP','PD','PP']
    rawDataFile['Delivery Terms'] = np.select(conditions,choices,default=np.nan)

#Service Level
    rawDataFile['Service Level'] = [SLDict.get(x,'') if x != np.nan else '' for x in rawDataFile['Service Level Code']]

#Load ID Date [GMT]
    rawDataFile['gtmDeltaOrg'] = [gmtDict.get(x,'') if x != np.nan else '' for x in rawDataFile['Orig. Airport']]
    rawDataFile['gtmDeltaOrg'] = rawDataFile['gtmDeltaOrg'].fillna(0).replace('', 0).astype(float)
    rawDataFile['gtmSignOrg'] = ['plus' if x >= 0 else 'minus' for x in rawDataFile['gtmDeltaOrg']]
    rawDataFile['gtmDeltaOrg'] = [ abs(int(x * 24 * 3600)) for x in  rawDataFile['gtmDeltaOrg']]
    rawDataFile['gtmDeltaOrg']  = [time(x//3600, (x%3600)//60, x%60) if x != 0 else time(0) for x in rawDataFile['gtmDeltaOrg']]
    rawDataFile['gtmDeltaOrg'] = pd.to_timedelta(rawDataFile['gtmDeltaOrg'].astype(str))
    rawDataFile['Load ID Date [GMT]'] = np.where(
            rawDataFile['gtmSignOrg'] == 'plus',
            rawDataFile['Event (ADD - %)'] + rawDataFile['gtmDeltaOrg'],
            rawDataFile['Event (ADD - %)'] - rawDataFile['gtmDeltaOrg']                                
                                                )

#On Hand Date [GMT Time]
    rawDataFile['Actual Pickup'] = pd.to_datetime(rawDataFile['Actual Pickup'])
    rawDataFile['On Hand Date [GMT Time]'] = np.where(
            rawDataFile['gtmSignOrg'] == 'plus',
            rawDataFile['Actual Pickup'] + rawDataFile['gtmDeltaOrg'],
            rawDataFile['Actual Pickup'] - rawDataFile['gtmDeltaOrg']                                
                                                    )

# Delivery to  Consignee [Local Time]
    rawDataFile['Delivery to  Consignee [Local Time]'] = np.where(rawDataFile['Delivery Terms'].isin(['DP','PP']),
                        rawDataFile['Event (Z70 - %)'],
                        rawDataFile['Actual Cartage Delivery']            
                                                                )                        

#Delivery to Consignee [GMT Time]
    rawDataFile['gtmDeltaDest'] = [gmtDict.get(x,'') if x != np.nan else '' for x in rawDataFile['Dest. Airport']]
    rawDataFile['gtmDeltaDest'] = rawDataFile['gtmDeltaDest'].fillna(0).replace('', 0).astype(float)
    rawDataFile['gtmSignDest'] = ['plus' if x >= 0 else 'minus' for x in rawDataFile['gtmDeltaDest']]
    rawDataFile['gtmDeltaDest'] = [ abs(int(x * 24 * 3600)) for x in  rawDataFile['gtmDeltaDest']]
    rawDataFile['gtmDeltaDest']  = [time(x//3600, (x%3600)//60, x%60) if x != 0 else time(0) for x in rawDataFile['gtmDeltaDest']]
    rawDataFile['gtmDeltaDest'] = pd.to_timedelta(rawDataFile['gtmDeltaDest'].astype(str))
    rawDataFile['Delivery to  Consignee [Local Time]'] = pd.to_datetime(rawDataFile['Delivery to  Consignee [Local Time]'])
    rawDataFile['Delivery to Consignee [GMT Time]'] = np.where(
            rawDataFile['gtmSignDest'] == 'plus',
            rawDataFile['Delivery to  Consignee [Local Time]'] + rawDataFile['gtmDeltaDest'],
            rawDataFile['Delivery to  Consignee [Local Time]'] - rawDataFile['gtmDeltaDest']   
                                                            )

# Transit Time SLA
    conditions = [ 
    (rawDataFile['Service Level'] != 'EM') & (rawDataFile['Service Level'] != 'RO') & \
    (rawDataFile['Service Level'] != 'PR'), #-
    (rawDataFile['House Ref'].str[0:3] =='LAX') & (rawDataFile['Service Level'] == 'EM') & \
    (rawDataFile['Consignor City'].isin(['LEHI','BOISE','EL PASO','RIO RANCHO','FORT COLLINS','COLORADO SPINGS'])), #72
    (rawDataFile['Service Level'] == 'EM'), #48
    (rawDataFile['Service Level'] =='PR') & (pd.to_timedelta(rawDataFile['Actual Pickup'].dt.strftime('%H:%M:%S').astype(str)) < datetime.timedelta(hours=18)),  #3
    (rawDataFile['Service Level'] =='PR') & (pd.to_timedelta(rawDataFile['Actual Pickup'].dt.strftime('%H:%M:%S').astype(str)) >= datetime.timedelta(hours=18)),  #4
    (rawDataFile['Service Level'] =='RO') & (pd.to_timedelta(rawDataFile['Actual Pickup'].dt.strftime('%H:%M:%S').astype(str)) < datetime.timedelta(hours=18,minutes=14)),  #5
    (rawDataFile['Service Level'] =='RO') & (pd.to_timedelta(rawDataFile['Actual Pickup'].dt.strftime('%H:%M:%S').astype(str)) >= datetime.timedelta(hours=18,minutes=14)),  #6
    ]
    choices = [np.nan,72,48,3,4,5,6]
    rawDataFile['Transit Time SLA'] = np.select(conditions,choices,default=np.nan)

# Due Date [Local]
    rawDataFile['Due Date [Local]'] = \
    np.where(
        (rawDataFile['Actual Pickup'].notnull()) & (~rawDataFile['Service Level'].isin(['EM','PR','RO'])),
        np.datetime64('NaT'),
        np.where(rawDataFile['Service Level'].isin(['PR','RO']),
        nextBusinessDay(rawDataFile['Actual Pickup'],rawDataFile['Transit Time SLA'],[pd.to_datetime(listsFileHolidays[x].values.reshape(-1)).dropna().tolist() for x in rawDataFile['Dest Ctry']]) + pd.Timedelta('23 hour 59 min 59 s'),
            np.where(rawDataFile['Service Level'].str == 'EM',
                (rawDataFile['Load ID Date [GMT]'] + pd.to_timedelta((rawDataFile['Transit Time SLA']/24),unit='D')),
                np.datetime64('NaT'))     
                )
            )

# Due Date [GMT]
    rawDataFile['Due Date [GMT]'] = np.where(rawDataFile['gtmSignDest'] == 'plus',
            rawDataFile['Due Date [Local]'] + rawDataFile['gtmDeltaDest'],
            rawDataFile['Due Date [Local]'] - rawDataFile['gtmDeltaDest']   
                                                            )

# Actual Transit Time
    rawDataFile['Actual Transit Time'] = np.where((rawDataFile['Service Level'] =='EM') & \
    (rawDataFile['Load ID Date [GMT]'].notnull()) & (rawDataFile['Delivery to Consignee [GMT Time]'].notnull()),
        ((rawDataFile['Delivery to Consignee [GMT Time]']-rawDataFile['Load ID Date [GMT]'])/pd.Timedelta('1 hour'))*24,
        np.where(
        (rawDataFile['Actual Pickup'].notnull()) | (rawDataFile['Delivery to  Consignee [Local Time]'].notnull()),
        countBusinessDays(rawDataFile['Actual Pickup'],rawDataFile['Delivery to  Consignee [Local Time]'],holidays=listsFileHolidays['US'].dropna().values.astype('datetime64[D]'))-1,
        np.nan
        ))

# Hours/ Days Late
    rawDataFile['Hours/ Days Late'] = rawDataFile['Actual Transit Time'] - rawDataFile['Transit Time SLA']

# On Time
    conditions = [ 
    (rawDataFile['Actual Transit Time'].isna()),  #In Transit
    ((rawDataFile['Hours/ Days Late'].isna()) | (rawDataFile['Hours/ Days Late'] <= 0)) #Yes
     ]

    choices = ['In Transit','Yes']

    rawDataFile['On Time'] = np.select(conditions,choices,default='No')

# Region
    rawDataFile['Region'] = rawDataFile['Orig. Region'] + '-' + rawDataFile['Dest. Region']

# Country-to-Country
    rawDataFile['Country-to-Country'] = rawDataFile['Origin Ctry'] + '-' + rawDataFile['Dest Ctry']

# Airport-to-Airport
    rawDataFile['Airport-to-Airport'] = rawDataFile['Orig. Airport'] + '-' + rawDataFile['Dest. Airport']

# Weight Break
    conditions = [ 
        (rawDataFile['Chargeable'] < 45), 
        ((rawDataFile['Chargeable'] >= 45) & (rawDataFile['Chargeable'] < 100)),
        ((rawDataFile['Chargeable'] >= 100) & (rawDataFile['Chargeable'] < 250)),
        ((rawDataFile['Chargeable'] >= 250) & (rawDataFile['Chargeable'] < 500)),
        ((rawDataFile['Chargeable'] >= 500) & (rawDataFile['Chargeable'] < 1000)),
        ((rawDataFile['Chargeable'] >= 1000) & (rawDataFile['Chargeable'] < 5000)),
        ((rawDataFile['Chargeable'] > 5000)),
        ]

    choices = ['< 45','> 45','> 100','> 250','> 500','> 1000','> 5000']

    rawDataFile['Weight Break'] = np.select(conditions,choices,default='')

# PU - ATD
    rawDataFile['PU - ATD'] = (rawDataFile['First Leg ATD'] - rawDataFile['Actual Pickup'])/pd.Timedelta('1 hour')

# ATD - ATA
    rawDataFile['ATD - ATA'] = (rawDataFile['Last Leg ATA'] - rawDataFile['First Leg ATD'])/pd.Timedelta('1 hour')

# ATA - DEL
    rawDataFile['ATA - DEL'] = np.where(rawDataFile['Service Level'] =='EM',
        (rawDataFile['Delivery to  Consignee [Local Time]'] - rawDataFile['Last Leg ATA'])/pd.Timedelta('1 hour'),
        countBusinessDays(rawDataFile['Delivery to  Consignee [Local Time]'],rawDataFile['Last Leg ATA'])*24
        )

# Total Transit Hours
    rawDataFile['Total Transit Hours'] = np.where(rawDataFile['Service Level'] =='EM',
                                                rawDataFile['Actual Transit Time'],
                                                rawDataFile['Actual Transit Time']*24)

# Pick-Up Q,m,W,d
    rawDataFile['Quarter (Pick-Up)'] = rawDataFile['Actual Pickup'].dt.to_period('Q').dt.strftime('Q%q-%Y')
    rawDataFile['Month (Pick-Up)'] = rawDataFile['Actual Pickup'].dt.strftime('%y-%m').astype(str)
    rawDataFile['Week (Pick-Up)'] = rawDataFile['Actual Pickup'].dt.strftime('%W').astype(str)
    rawDataFile['Day (Pick-Up)'] = rawDataFile['Actual Pickup'].dt.strftime('%A').astype(str)

# Delay Owner
    conditions = [ 
    (rawDataFile['On Time'] == 'Yes'),                           #-
    (rawDataFile['On Time'] == 'In Transit'),                    #-
    (rawDataFile['ATA - DEL'] < 0),                              #Origin
    ((rawDataFile['ATA - DEL'] <= 12) & (rawDataFile['Service Level'] == 'EM')), #Origin
    ((rawDataFile['ATA - DEL'] <= 24) & (rawDataFile['Service Level'] == 'PR')), #Origin
    ((rawDataFile['ATA - DEL'] <= 48) & (rawDataFile['Service Level'] == 'RO')), #Origin
     ]

    choices = ['-','-','Origin','Origin','Origin','Origin']

    rawDataFile['Delay Owner'] = np.select(conditions,choices,default='Destination')

# Delay Code
    rawDataFile['Delay Code'] = np.where(rawDataFile['Delay Owner'] == '-',
                            '',
                            np.where(rawDataFile['Delay Owner'] == 'Origin',
                                     rawDataFile['Shipment - Custom Field 01'].str.upper().str[:3],
                                     rawDataFile['Shipment - Custom Field 02'].str.upper().str[:3]
                                    ))

# Delay Description
    rawDataFile['Delay Description'] = [delayDescriptionDict.get(x,'') if x != '' else '' for x in rawDataFile['Delay Code']]

# Controllable
    rawDataFile['Controllable'] = [delayControllableDict.get(x,'') if x != '' else '' for x in rawDataFile['Delay Code']]

# Nett
    rawDataFile['Nett'] = np.where(rawDataFile['On Time'] == 'Yes',
           1,
           np.where(((rawDataFile['On Time'] == 'Yes') &( rawDataFile['Controllable'] == 'N')),
                1,
                np.nan))

# Count
    rawDataFile['Count'] = 1

# Gross
    rawDataFile['Gross'] = np.where(rawDataFile['On Time'] == 'Yes',
                    1,
                    np.nan)

# Rename columns names
    rawDataFile = rawDataFile.rename(columns={i:j for i,j in zip(rawDataColumns,finalDataColumns)})

# All Columns List in order
    allColumnList = [
    "HAWB",
    "Shipper Name",
    "Shipper City",
    "Consignee Name",
    "Consignee City",
    "Orig. Country",
    "Dest. Country",
    "Inco Terms",
    "Pcs",
    "Act. Weight",
    "Chg. Weight",
    "Load ID Date [Local]",
    "On Hand Date [Local Time]",
    "ETD First Load",
    "ETA Last Disch",
    "Actual Time of Departure [Local]",
    "Actual Time of Arrival [Local]",
    "MAWB",
    "Airline",
    "Orig. Region",
    "Orig. Airport",
    "Dest. Region",
    "Dest. Airport",
    "Delivery Terms",
    "Service Level",
    "Load ID Date [GMT]",
    "On Hand Date [GMT Time]",
    "Delivery to  Consignee [Local Time]",
    "Delivery to Consignee [GMT Time]",
    "Transit Time SLA",
    "Due Date [Local]",
    "Due Date [GMT]",
    "Actual Transit Time",
    "Hours/ Days Late",
    "On Time",
    "Region",
    "Country-to-Country",
    "Airport-to-Airport",
    "Weight Break",
    "PU - ATD",
    "ATD - ATA",
    "ATA - DEL",
    "Total Transit Hours",
    "Quarter (Pick-Up)",
    "Month (Pick-Up)",
    "Week (Pick-Up)",
    "Day (Pick-Up)",
    "Delay Code",
    "Delay Description",
    "Controllable",
    "Nett",
    "Gross",
    'Count'
    ]


# Return
    return rawDataFile[allColumnList]




def startProgram():

    start = timer()

    rawDataPath = rawDataEntry.get()
    listsFilePath = listsEntry.get()
    mappingFilePath = mappingEntry.get()
    fileDirectory = newPathEntry.get()

    break1 = (timer()-start)
    print(f'Ścieżki pobrane w {break1} sekund')

    rawDataFile, prefixDict ,countryRegionDict, SLDict,gmtDict,listsFileHolidays,delayDescriptionDict,delayControllableDict,rawDataColumns,finalDataColumns = loadingFiles(rawDataPath,listsFilePath,mappingFilePath)

    break2 = (timer()-start)
    print(f'Pliki załadowane w {break2} sekund')

    rawDataFile = prepareRawDataASML(rawDataFile,prefixDict,countryRegionDict,SLDict,gmtDict,listsFileHolidays,delayDescriptionDict,delayControllableDict,rawDataColumns,finalDataColumns)

    break3 = (timer()-start)
    print(f'Plik obrobiony w {break3} sekund')

    fileDirectory= create_folder(fileDirectory)
    rawDataFile.to_excel(f'{fileDirectory}Raw_Data.xlsx')

    break4 = (timer()-start)
    print(f'Całkowity czas działania {break4} sekund')

    end = timer()
    result = end-start
    if result > 60:
        timeMessage = f'Uff {float("{:.2f}".format(result))} second! It took a while but your report is completed'
    else:
        timeMessage = f'It was fast! {float("{:.2f}".format(result))} seconds! Is it a new record? Enjoy your report!'

    messagebox.showinfo(f'Report is ready!',f'{timeMessage}\nFile is saved in below directory:\n {fileDirectory}Raw_Data.xlsx')

root = Tk()
root.title('ASML Report Creator')
root.geometry('740x140')

# ============vvvvvvvvvv============ENTRY AND BUTTONS FOR OPENING FOLDERS AND FILES============vvvvvvvvvv============

#Entry for the main workbook path
listsEntry = Entry(root, width=90,borderwidth=2)
listsEntry.grid(column=1,columnspan=2,row=1)
listsButton = Button(root, text='Choose lists file',command=GetListFilePath,width=25).grid(column=0,row=1)

#Entry for the path to raw data
rawDataEntry = Entry(root, width=90,borderwidth=2)
rawDataEntry.grid(column=1,columnspan=2,row=0)
rawDataButton = Button(root, text='Choose raw data file',command=GetRawDataFilePath,width=25).grid(column=0,row=0)

#Entry for the path for output file
newPathEntry = Entry(root, width=90,borderwidth=2)
newPathEntry.grid(column=1,columnspan=2,row=3)
newPathButton = Button(root, text='Choose directory for your report',command=GetNewFilePath,width=25).grid(column=0,row=3)

#Entry for the path to mapping file
mappingEntry = Entry(root, width=90,borderwidth=2)
mappingEntry.grid(column=1,columnspan=2,row=4)
mappingButton = Button(root, text='Choose your Mapping file',command=GetMappingFilePath,width=25).grid(column=0,row=4)


# ============^^^^^^^^^^============ENTRY AND BUTTONS FOR OPENING FOLDERS AND FILES============^^^^^^^^^^============

#Start Program button
StartButton = Button(root,text='Start',command=startProgram,height=1,width=10).grid(column=0,columnspan=3)
root.mainloop()



