import pandas as pd
import csv
import time



trackpath = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup_Tracker\\Data\\tracked_setups.csv"
untrackpath = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup_Tracker\\Data\\untracked_setups.csv"
wippath = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup_Tracker\\Data\\setups_in_progress.csv"

#trackpath = ".\\Data\\tracked_setups.csv"
#wippath = ".\\Data\\setups_in_progress.csv"
#untrackpath = ".\\Data\\untracked_setups.csv"

metricspath = "S:\\Metrics\\5 - Grinding\\setup_tracking.csv"
#metricspath = ".\\Data\\setup_tracking.csv"

def writeSetupCSV(path, source, setup_list=[]):
    while True:
        try:
            with open(path, 'a', newline= '') as setupcsv:
                trackwriter = csv.writer(setupcsv, delimiter=',', quoting=csv.QUOTE_MINIMAL)        
                trackwriter.writerow(setup_list)
                print(source + " file write successful.")
                break
        except PermissionError as pe:
            print(pe, "File is open in excel. Data will save once the file is closed.")
            time.sleep(1)
            continue    

def writeWip(sd=['00000','nomach','no Opr']):
    now = pd.Timestamp.now()
    today = pd.Timestamp.date(now)
    when = pd.Timestamp.time(now)
    sd.extend([today,when])
    writeSetupCSV(wippath,'WIP',sd)

def findCompWip(jobnum):
    df = pd.read_csv(wippath)    
    dfilter = df['Job Number'].astype(str) == jobnum    
    if df[dfilter].values.size > 0:
        found = df[dfilter].values
        newdf = df[(df['Job Number'].astype(str) != jobnum)]
        while True:
            try:
                newdf.to_csv(wippath, sep=',', index=False)
                print("WIP file write successful.")
                break
            except PermissionError as pe:
                print(pe, "File is open in excel. Data will save once the file is closed.")
                time.sleep(1)
                continue
        return found[0]
    else:        
        return None

def writeCompleted(jobnum="",comment="",breaktime=False,lunch=False):
    complete_data = findCompWip(jobnum)   
    
    if complete_data.size == 0:
        print('No data to track.')
        return
    now = pd.Timestamp.now()
    today = pd.Timestamp.date(now)
    finish = pd.Timestamp.time(now)
    start = pd.to_datetime(complete_data[3] + 'T' + complete_data[4])    
    total_hours = pd.Timedelta(now - start).total_seconds()
    breaktime = bool(breaktime)
    lunch = bool(lunch)
    if breaktime:
        total_hours = total_hours - 900
 
    if lunch:
        total_hours = total_hours - 1800
  
    total_hours = total_hours/3600
    ml = complete_data.tolist() #ml is mylist
    ml.extend([finish,total_hours,comment,breaktime,lunch])   
    #extract job number, machine, operator, finish date, total hours to send to writeSetupCSV for metrics usage
    metrics_list = [ml[0],ml[1],ml[2],today,total_hours]    
    
    writeSetupCSV(trackpath,'Setup Tracking',ml)
    try:
        writeSetupCSV(metricspath,'Setup Metrics',metrics_list)
    except:
        print('Unable to write to Setup Metrics')
    
if __name__ == "__main__": 
    #writeWip(['12345','VanDamme','Engineering Test'])
    writeCompleted('93530a')