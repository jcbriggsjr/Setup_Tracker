import openpyxl

path = "G:\\3 - Production Departments\\11- Scheduling\\Grinding\\Setup_Tracking.xlsx"
sublist = [0,0]

def main():
    return readwipSheet()

def readwipSheet():
    wb = openpyxl.load_workbook(path,read_only=True)
    sh = wb["Setups In Progress"]
    maxdatarow = 2 #initialize to first line, in case of empty list
    for row in sh.iter_rows(min_row=2, max_col=2, max_row=sh.max_row):
        for cell in row:
            if cell.value != None:
                maxdatarow = cell.row
    wiplist = [[None]*2 for i in range(maxdatarow-1)]
    for row in sh.iter_rows(min_row=2, max_col=2, max_row=maxdatarow):            
        for cell in row:
            if cell.value != None and cell.column == 1:
                wiplist[cell.row-2][0] = cell.value              
            if type(cell.value) == type("string") and cell.column == 2:
                wiplist[cell.row-2][1] = cell.value
    return wiplist
            
if __name__ == "__main__":
    print(main())