#v 2.1.1
# tool_setter.py contains program pull/save function
from datetime import *
from tkinter import *
from updateoptionmenu import *
import tkinter.messagebox as tkMessageBox
import gemployees
import setupwip
import writesetupwip
import pyodbc
import tool_setter

master = Tk()
master.title("Setup Tracking")

frame = Frame(master, bg="pale turquoise")
frame.grid()

popup = Tk()
popup.wm_withdraw()

def update():
    wip.options = setupwip.readWIP() #csv 
    test = wip.options[0][0]        
    if 'None' in wip.options[0]:
        wip.om_variable.set('None')
    elif not wip.om_variable.get() in str(wip.options):
        wip.om_variable.set(wip.options[0])
        
    wip.update_option_menu()
    master.after(10000,update)
    
def displayError():
    #message at x:200,y:200
    popup.geometry("1x1+200+200")#remember its .geometry("WidthxHeight(+or-)X(+or-)Y")
    tkMessageBox.showerror(title="Bad Job Number",message="Job # not found. Please enter correct job #.",parent=popup)
                           
def displayProgError(status):
    #message at x:200,y:200
    popup.geometry("1x1+200+200")#remember its .geometry("WidthxHeight(+or-)X(+or-)Y")
    tkMessageBox.showerror(title="Program Sending Status",message=status,parent=popup)    
    
def displayProgSuccess(status):
    #message at x:200,y:200
    popup.geometry("1x1+200+200")#remember its .geometry("WidthxHeight(+or-)X(+or-)Y")
    tkMessageBox.showinfo(title="Program Sending Status",message=status,parent=popup)    

def displayToolError(status):
    #message at x:200,y:200
    popup.geometry("1x1+200+200")#remember its .geometry("WidthxHeight(+or-)X(+or-)Y")
    tkMessageBox.showerror(title="Tool Data Sending Status",message=status,parent=popup)

def displayToolSuccess(status):
    #message at x:200,y:200
    popup.geometry("1x1+200+200")#remember its .geometry("WidthxHeight(+or-)X(+or-)Y")
    tkMessageBox.showinfo(title="Tool Data Sending Status",message=status,parent=popup)
    
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
    writesetupwip.writeWip([jobnum,machine,operator]) #csv version
    outputstring = str(machine) + " " + str(operator) + " " + str(datetime.today().time().strftime("%H:%M:%S")) + " job # " + str(jobnum)
    makepopup(outputstring)
    wip.options = setupwip.readWIP() #csv version
    wip.update_option_menu()
    wip.om_variable.set(wip.options[0])
    jobentry.delete(0, 'end')
    
def writetoFinish():
    long_setup_reason = reason_entry.get()
    completed_job = filterJobNumber(wip.om_variable.get())
    isbreak = breaktime.get()
    islunch = lunch.get()
    writesetupwip.writeCompleted(completed_job, long_setup_reason, isbreak, islunch) #csv version
    wip.options = setupwip.readWIP()
    wip.update_option_menu()
    wip.om_variable.set(wip.options[0])
    reason_entry.delete(0,'end')
    breaktime.set(0)
    lunch.set(0)

def filterJobNumber(rawdata):
    endof = rawdata.find(' ')
    job_number = rawdata[1:endof]
    if "'" in job_number:
        job_number = job_number.split("'")[1]
    return job_number

def filterMachine(rawdata):
    start = rawdata.find(' ')
    machine = rawdata[start:]
    if "'" in machine:
        machine = machine.split("'")[1]
    print(machine)
    return machine

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

def sendProg():
    job_num = filterJobNumber(wip.om_variable.get())    
    machine = filterMachine(wip.om_variable.get())
    if machine not in communicable_list:
        status = "Can't send data to machine yet. Contact your programmer."
        displayProgError(status)
    else:
        status = tool_setter.process_program_request(job_num, machine)
        if status[:5] != 'Sendi':
            displayProgError(status)
        else:
            displayProgSuccess(status)
    
def sendTool():
    job_num = filterJobNumber(wip.om_variable.get())
    machine = filterMachine(wip.om_variable.get())
    if machine not in communicable_list:
        status = "Can't send data to machine yet. Contact your programmer."
        displayToolError(status)
    else:
        status = tool_setter.process_tool_request(job_num, machine)
        if status[:5] != 'Tool ':
            displayToolError(status)
        else:
            displayToolSuccess(status)

def sendProbe():
    job_num = filterJobNumber(wip.om_variable.get())
    machine = filterMachine(wip.om_variable.get())
    if machine not in communicable_list:
        status = "Can't send data to machine yet. Contact your programmer."
        displayProgError(status)
    else:
        status = tool_setter.process_strikeprobe_request(job_num, machine, 'Probing')
        if status[:5] != 'Probi':
            displayProgError(status)
        else:
            displayProgSuccess(status)
            
def sendStrike():
    job_num = filterJobNumber(wip.om_variable.get())
    machine = filterMachine(wip.om_variable.get())
    if machine not in communicable_list:
        status = "Can't send data to machine yet. Contact your programmer."
        displayProgError(status)
    else:
        status = tool_setter.process_strikeprobe_request(job_num, machine, 'Striking')
        print(status)
        if status[:5] != 'Strik':
            displayProgError(status)
        else:
            displayProgSuccess(status)
    

def close_window():
    master.destroy()
    master.quit()
    
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
    'Dutch',    
    'Gambit',
    'LittleBrother',
    'Rocky',    
    'VanDamme']
# =============================================================================
# auxlist = [
#     'Belly Saw',
#     'Beveler',
#     'Blanchard',
#     'Bothner',
#     'Dama Llama',
#     'Mini Drill',
#     'Opposed Drill',
#     'Polish/De-block',
#     'Segway',
#     'Slicer Dicer']
# =============================================================================

communicable_list = ['Chuck', 'VanDamme', 'Bruce', 'Dutch']
combinedlist = vmclist #+ auxlist #auxlist commented out as it likely won't get used 12/31/2020
combinedlist.sort()

machine_dropdown = Dropdown(frame,combinedlist,jobentry.get())
machine_dropdown.om_variable.set(combinedlist[0])
machine_dropdown.om.grid(row=1,column=2)

#create label over setup operator list
operatorlabel = Label(frame, text="Setup Operator", bg="pale turquoise", height=2)
operatorlabel.grid(row=0,column=3)

#create setup operator name list
oplist = gemployees.readEmployees()
op_dropdown = Dropdown(frame,oplist,jobentry.get())
op_dropdown.om_variable.set(oplist[0])
op_dropdown.om.grid(row=1, column=3)

#create button to log start of setup
startbutton = Button(frame, text = "Start", width=20, command=writetoWIP, bg="light green", activebackground="light blue")
startbutton.grid(row=1,column=4)

#create label in front of wip dropdown list
Label(frame, text='Select In-progress Job #', bg="pale turquoise").grid(row=2)

#create dropdown to select wip jobs to complete. Display job# and machine.
wipjoblist = setupwip.readWIP()
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
Label(frame, text='Enter reason for setup longer than 1 hour:', bg="pale turquoise").grid(row=3)
reason_entry = Entry(frame)
reason_entry.grid(row=3,column=1)

#create button to send Programs to selected machine
send_prog_button = Button(frame, text = "Send Program", width=20, command=sendProg, bg="light green", activebackground="light blue")
send_prog_button.grid(row=4,column=3)

send_tool_button = Button(frame, text="Send Tools", width=20, command=sendTool, bg="light green", activebackground="light blue")
send_tool_button.grid(row=4,column=4)

send_probe_button = Button(frame, text = "Send Probing", width=20, command=sendProbe, bg="light green", activebackground="light blue")
send_probe_button.grid(row=5,column=4,padx=5,pady=10)

# send strike button is commented out for current release. probing seems to work well, striking needs more testing 4-7-2020
send_strike_button = Button(frame, text = "Send Strike",width=20, command=sendStrike, bg="light green", activebackground="light blue")
send_strike_button.grid(row=5,column=3,padx=5,pady=10)

breaktime = IntVar()
breakbox = Checkbutton(frame, variable=breaktime, onvalue=True, offvalue=False, text ='Break', takefocus = 0, bg='pale turquoise')
breakbox.grid(row=2,column=2)

lunch = IntVar()
lunchbox = Checkbutton(frame, variable=lunch, onvalue=True, offvalue=False, text='Lunch', takefocus = 0, bg='pale turquoise')
lunchbox.grid(row=2,column=3)

Button(master, text="Quit", width=30, command=close_window).grid()

master.after(10000,update)
master.mainloop()