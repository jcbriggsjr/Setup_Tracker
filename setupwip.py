import pandas as pd

path = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup Tracker\\Data\\setups_in_progress.csv"
#path = ".//Data//setups_in_progress.csv" #testing path
sublist = [0,0]

def main():
    return readWIP()

def readWIP():
    data = pd.read_csv(path)    
    wip = data.loc[:,'Job Number':'Machine'].values
    if wip.size == 0:
        wip = ['None']
    return wip    
    
        
if __name__ == "__main__":
    print(main())