import csv

path = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup Tracker\\Data\\grinding_employees.csv"

def main():
    return readEmployees()

def readEmployees():
    emplist = []
    
    with open(path, newline='\n') as f:
        reader = csv.reader(f)
        readlist = list(reader)
        
    for e in readlist:
        emplist.append(e[0])
    emplist.sort()
        
    return emplist
if __name__ == "__main__":
    print(main())



    