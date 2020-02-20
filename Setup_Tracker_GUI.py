from datetime import *
from tkinter import *
from updateoptionmenu import *
import tkinter.messagebox as tkMessageBox
import gemployees
import setupwip
import writesetupwip
import pyodbc

master = Tk()
master.title("Setup Tracking")

frame = Frame(master, bg="pale turquoise")
frame.grid()

popup = Tk()
popup.wm_withdraw()
def displayError():
    #message at x:200,y:200
    popup.geometry("1x1+200+200")#remember its .geometry("WidthxHeight(+or-)X(+or-)Y")
    tkMessageBox.showerror(title="Bad Job Number",message="Job # not found. Please enter correct job #.",parent=popup)

#center screen message
def makepopup(newmessage):
    popup.geometry("1x1+"+str(int(popup.winfo_screenwidth()/2))+"+"+str(int(popup.winfo_screenheight()/2)))
    tkMessageBox.showinfo(title="Start Confirmation", message=newmessage)

def writetoWIP():
    #write to WIP should store job #, machine, Operator, and time    
    machine = machine_dropdown.om_variable.get()
    operator = op_dropdown.om_variable.get()
    jobnum = jobentry.get()
    if not getData(jobnum):
        displayError()
        return
    writesetupwip.writewipSheet(jobnum, machine, operator)
    outputstring = str(machine) + " " + str(operator) + " " + str(datetime.today().time().strftime("%H:%M:%S")) + " job # " + str(jobnum)
    makepopup(outputstring)
    wip.options = setupwip.readwipSheet()
    wip.update_option_menu()
    wip.om_variable.set(wip.options[0])
    jobentry.delete(0, 'end')
    
def writetoFinish():
    long_setup_reason = reason_entry.get()
    completed_job = filterJobNumber(wip.om_variable.get())
    writesetupwip.writeCompletedSheet(completed_job, long_setup_reason)
    wip.options = setupwip.readwipSheet()
    wip.update_option_menu()
    wip.om_variable.set(wip.options[0])
    reason_entry.delete(0,'end')

def filterJobNumber(rawdata):
    rawdata = rawdata[2:]
    rawdata = rawdata.split("'")
    return rawdata[0]
    
def getData(jobnum):
    jobnum = "'" + jobnum + "'"
    jobpresent = True #boolean initialization
    count = 0 #counter to determine if we have data from the cursor
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=SVR-APP\\SQLEXPRESS;'
                          'Database=QssCatiJobTrack;'
                          'Uid=sa;'
                          'Pwd=$ql4Q$$4C@T1;'
                          'Trusted_Connect=yes;') #connection string, connects to DB
    
    cursor = conn.cursor() #creates cursor object
    #the query below 
    cursor.execute("SELECT Jobs.JobNum "
                   "FROM QssCatiJobTrack.dbo.Jobs "
                   "WHERE Jobs.JobNum= " + jobnum + " ") #query matches the jobnumber exactly. Will return nothing if None

    for row in cursor:
        count = count +1
    
    if count > 0:
        jobpresent = True
    else:
        jobpresent = False
    cursor.close()
    conn.close()    
    return jobpresent
    
#create job # input box
Label(frame, text='Input Job #', bg="pale turquoise").grid(row=1)
jobentry = Entry(frame)
jobentry.grid(row=1,column=1)

#create label for dropdown list to select machine
droplabel = Label(frame, text="Setup Machine", bg="pale turquoise", height=2)
droplabel.grid(row=0,column=2)

#create dropdown list to select machine
vmclist = [
    'Bruce',
    'Chuck',    
    'Gambit',
    'Little Bro',
    'Rocky',    
    'Van Damme']
auxlist = [
    'Belly Saw',
    'Beveler',
    'Blanchard',
    'Bothner',
    'Dama Llama',
    'Mini Drill',
    'Opposed Drill',
    'Polish/De-block',
    'Segway',
    'Slicer Dicer']
combinedlist = vmclist + auxlist
combinedlist.sort()

machine_dropdown = Dropdown(frame,combinedlist,jobentry.get())
machine_dropdown.om_variable.set(combinedlist[0])
machine_dropdown.om.grid(row=1,column=2)

#create label over setup operator list
operatorlabel = Label(frame, text="Setup Operator", bg="pale turquoise", height=2)
operatorlabel.grid(row=0,column=3)

#create setup operator name list
oplist = gemployees.readSheet()
op_dropdown = Dropdown(frame,oplist,jobentry.get())
op_dropdown.om_variable.set(oplist[0])
op_dropdown.om.grid(row=1, column=3)

#create button to log start of setup
startbutton = Button(frame, text = "Start", width=20, command=writetoWIP, bg="light green", activebackground="light blue")
startbutton.grid(row=1,column=4)

#create label in front of wip dropdown list
Label(frame, text='Select Completed Job #', bg="pale turquoise").grid(row=2)

#create dropdown to select wip jobs to complete. Display job# and machine.
wipjoblist = setupwip.readwipSheet()
wip = Dropdown(frame, wipjoblist, jobentry.get())
wip.om.grid(row=2,column=1)
wip.om_variable.set(wipjoblist[0])

#create button to log completion of setup
finishbutton = Button(frame, text = "Complete", width=20, command=writetoFinish, bg="light green", activebackground="light blue")
finishbutton.grid(row=2,column=4)

#create label for long setup entry box
#reason_label = Label(frame, text="Reason for long setup ( >1 hour):", bg="pale turquoise", height=2)
#reason_label.grid(row=3, column=0)

#create entry box for long setup reason
Label(frame, text='Reason for long setup ( >1 hour) :', bg="pale turquoise").grid(row=3)
reason_entry = Entry(frame)
reason_entry.grid(row=3,column=1)

master.mainloop()