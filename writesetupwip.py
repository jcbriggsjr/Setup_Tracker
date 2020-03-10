import pandas as pd
import csv
import time

#excelpath = r".\\Data\\Setup_Tracking.xlsx"
excelpath = r"G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup Tracker\\Data\\Setup_Tracking.xlsx"
trackpath = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup Tracker\\Data\\tracked_setups.csv"
untrackpath = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup Tracker\\Data\\untracked_setups.csv"
wippath = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup Tracker\\Data\\setups_in_progress.csv"

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
        
def writeExcel(inpath, outpath):
    
    # get latest data from csv to put in excel
    df = pd.read_csv(inpath)
    
    # write out the new sheet. create a writer object using openpyxl in append mode    
    # try to open excel file and modify. will fail if file is open
    try:
        with pd.ExcelWriter(excelpath, engine='openpyxl', mode='a') as writer:
            
            # create workbook object
            workbook = writer.book
            
            # try to remove the existing sheet so we don't end up creating multiple copies
            try:
                workbook.remove(workbook['Tracked Setups'])
            except:
                print("Worksheet doesn't exist")
                
            # write csv data back into excel sheet
            finally:
                df.to_excel(writer, sheet_name='Tracked Setups', index=False)
        
        # save the workbook
        writer.save()
    except:
        print("Excel file is open. Will update when file is closed.")
    

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

def writeCompleted(jobnum="",comment=""):
    complete_data = findCompWip(jobnum)   
    
    if complete_data.size == 0:
        print('No data to track.')
        return
    now = pd.Timestamp.now()
    today = pd.Timestamp.date(now)
    finish = pd.Timestamp.time(now)
    start = pd.to_datetime(complete_data[3] + 'T' + complete_data[4])    
    total_hours = pd.Timedelta(now - start).total_seconds()/3600  
    ml = complete_data.tolist() #ml is mylist
    ml.extend([finish,total_hours,comment])   
    #extract job number, machine, operator, finish date, total hours to send to writeSetupCSV for metrics usage
    metrics_list = [ml[0],ml[1],ml[2],today,total_hours]    
    
    writeSetupCSV(trackpath,'Setup Tracking',ml)
    writeSetupCSV(metricspath,'Setup Metrics',metrics_list)
    
    # don't need to write data to excel. Separate excel file will be created to generate graphs
    #writeExcel(trackpath, excelpath)
    
if __name__ == "__main__": 
    #writeWip(['12345','VanDamme','Engineering Test'])
    writeCompleted('93530a')