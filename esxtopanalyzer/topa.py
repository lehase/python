import os
import re
import pandas as pd
from collections import OrderedDict
from datetime import date
from termcolor import colored
import colorama
from subprocess import Popen, PIPE
import time
import multiprocessing


def AnalyzeVMList(filepath):
    DsList = list()
    try:
        print(colored('Reading file {}'.format(filepath), 'green'))
        with open(filepath) as fp:  
            data=fp.read()
            wid_pattern=re.compile(r'World ID: (\d+)')
            vmname_pattern=re.compile(r'Display Name: (\S+)')
                        
            wid=wid_pattern.findall(data)
            vmname=vmname_pattern.findall(data)
            
            #Translate lists to series
            wid=pd.Series(wid)
            print(wid)
            vmname=pd.Series(vmname)
            print(vmname)
            
            WID_VM=pd.DataFrame()
            WID_VM['wid']=wid
            WID_VM['VM']=vmname
            print(colored('      Dataframe is:', 'green'))
            print(WID_VM)
            #WID_VM = pd.DataFrame.from_records(se)
            #WID_VM=WID_VM.assign(wid,  .values())
            WID_VM['wid']=WID_VM['wid'].astype('int64', copy=False)
    finally:
        fp.close()
    
    return WID_VM

def DatastoreAnalyse(filepath):
    DS_List=pd.DataFrame()
    print(colored('Reading file {}'.format(filepath), 'green'))
    DS_List = pd.read_fwf(filepath, header=0, usecols=[0, 3], skipinitialspace=True, memory_map=True) #, skiprows=[1])
    
    #print(DS_List)
    return DS_List

def PerfLoad(filepath):
    print(colored('Reading file {}'.format(filepath), 'green'))
    PerfomanceDF=pd.DataFrame()
    FilteredDF=pd.DataFrame()
    NormalizedDF=pd.DataFrame()
    ComposedDF=pd.DataFrame()
    PerfomanceDF=pd.read_csv(filepath, parse_dates=[1], infer_datetime_format=True, quotechar='"', memory_map = True)
    print(colored('Transposing Dataframe', 'green'))
    PerfomanceDF=PerfomanceDF.transpose()
    #Делаем первую строку - заголовком
    PerfomanceDF.columns=PerfomanceDF.iloc[0]
    PerfomanceDF=PerfomanceDF.drop(PerfomanceDF.index[0])
    print(colored('Apply filter to DF', 'green'))
    FilteredDF=PerfomanceDF.filter(like='Disk Per-Device-Per-World', axis=0)
    #Resetting index (Conter names go to data)
    FilteredDF.reset_index(drop=False, inplace=True)
    # Количество колонок len(FilteredDF.columns)
    # Имя колонки print(FilteredDF.columns[1])
    #Возврат пары знчение + время
    print(colored('Stretch Matrix to flat table', 'green'))
    for i in range(len(FilteredDF.columns)-1):
        NormalizedDF=FilteredDF.iloc[:,[0,i+1]]
        NormalizedDF.columns = ['Item', 'Value']
        NormalizedDF['Date']=FilteredDF.columns[i+1]
        ComposedDF=pd.concat([ComposedDF, NormalizedDF], ignore_index=True)

    
    
    #########################################
    #PostRocessing
    print(colored('Converting to DateTime', 'green'))
    ComposedDF['Date']=pd.to_datetime(ComposedDF['Date'], infer_datetime_format=True)
    print(colored('Splitting columns', 'green'))
    ComposedDF['Host']=ComposedDF['Item'].str.split('\\').str.get(2)
    ComposedDF['Parent']=ComposedDF['Item'].str.split('\\').str.get(3)
    ComposedDF['Counter']=ComposedDF['Item'].str.split('\\').str.get(4)
    ComposedDF['DeviceID']=ComposedDF['Parent'].str.split('\(').str.get(1).str.replace('\)','')
    ComposedDF['WorldID']=ComposedDF['DeviceID'].str.split('-').str.get(1)
    ComposedDF['WorldID']=ComposedDF['WorldID'].astype('int64', copy=False)
    ComposedDF['DeviceID']=ComposedDF['DeviceID'].str.split('-').str.get(0)
    #Join with VM Names and Storage Names
    print(colored('Join with VM Names and Storage Names', 'green'))
    ComposedDF=ComposedDF.join(WID_VM.set_index('wid'),on='WorldID')
    ComposedDF=ComposedDF.join(DS_List.set_index('Device Name'),on='DeviceID')
   # print(ComposedDF)
    print(colored('Drop huge columns', 'green'))
    ComposedDF.drop(columns=['Item', 'Parent', 'DeviceID', 'WorldID'], inplace=True)

    
    #print(FilteredDF)
    
    #print(colored('Saving to excel', 'green'))
    #writer = pd.ExcelWriter(os.path.dirname(os.path.abspath(__file__)) +"""\\output.xlsx""")
    #PerfomanceDF.to_excel(writer,'Sheet1')
    #ComposedDF.to_excel(writer,'Sheet2')
    #writer.save()
    return ComposedDF
def GetHostsConfig(filepath):
    #ESXList = list()
    print(colored('Reading file {}'.format(filepath), 'green'))
    ESXList = pd.read_csv(filepath,  skipinitialspace=True, memory_map=True, skip_blank_lines=True, header=0, delimiter = ';') #, skiprows=[1])
    
    print(ESXList)
    return ESXList
    
if __name__ == '__main__':
    colorama.init()
    #Disable SettingWithCopyWarning
    pd.options.mode.chained_assignment = None
    config_folder= os.path.dirname(os.path.abspath(__file__))
    datafolder = os.path.dirname(os.path.abspath(__file__)) + """\\data\\"""
    DatastoreList = datafolder + 'dslist.txt'
    VMList = datafolder + 'vms.txt'
    PerfomanceCSV=datafolder + 'esxc02.csv'
    ESX_List_file=config_folder + '\\esxlist.txt'
    
    #Detect CPU count
    CPU_count=multiprocessing.cpu_count()
    print('We have {} CPUs'.format(CPU_count))
    

    #Read config with ESXi hosts
    ESXList=GetHostsConfig(ESX_List_file)
    for i in range(len(ESXList)):
        host = ESXList.iloc[i,0]
        login = ESXList.iloc[i,1]
        password = ESXList.iloc[i,2]
        print(host, login, password)
    
    
    
    #print(ESXList.columns)
    
    #WID_VM=AnalyzeVMList(VMList)
    #DS_List=DatastoreAnalyse(DatastoreList)
    #PerfomanceDF=PerfLoad(PerfomanceCSV)
    

    print(colored('END','red'))
