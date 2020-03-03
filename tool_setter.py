# -*- coding: utf-8 -*-
"""
Created on Fri Feb 21 16:09:13 2020

@author: JBriggs
"""

import socket
import glob
import os

path = 'G:\\3 - Production Departments\\4 - Grinding\\9 - VMCs\\4 - Programs & Software\\10-Programs_and_Tools_to_Transfer\\'
 
machineip = {'Bruce':'192.168.84.201',
            'Chuck':'192.168.84.200',
            'VanDamme':'192.168.84.203'}


brucepp = path + 'Bruce\\Programs\\'
chuckpp = path + 'Chuck\\Programs\\'
vandammepp = path + 'VanDamme\\Programs\\'

brucetp = path + 'Bruce\\Tools\\'
chucktp = path + 'Chuck\\Tools\\'
vandammetp = path + 'VanDamme\\Tools\\'

bruceap = path + 'Bruce\\ATC_data\\'
chuckap = path + 'Chuck\\ATC_data\\'
vandammeap = path + 'VanDamme\\ATC_data\\'

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
BUFFER_SIZE = 1024

def connect_to_machine(machine):
    TCP_IP = machineip[machine]
    TCP_PORT = 10000    
    s.connect((TCP_IP, TCP_PORT))
    

def disconnect_machine():
    s.close

def set_paths(machine):
    tp = {
            'Bruce': brucetp,
            'Chuck': chucktp,
            'VanDamme': vandammetp}
    pp = {  
            'Bruce': brucepp,
            'Chuck': chuckpp,
            'VanDamme': vandammepp}
    
    ap = {
            'Bruce': bruceap,
            'Chuck': chuckap,
            'VanDamme': vandammeap
            }
    return pp.get(machine,'No program path'), tp.get(machine,'No tool path'), ap.get(machine,'No ATC path')

def find_save_path(customer):
    search_dir = 'G:\\3 - Production Departments\\4 - Grinding\\9 - VMCs\\4 - Programs & Software\\2 - Programs & MC Files - Haas'    
    dir_list = os.listdir(search_dir)
    folder = [s for s in dir_list if customer.lower() in s.lower()]
    if len(folder) > 0:
        return search_dir + '\\' + folder[0] + '\\'
    return search_dir + '\\' + "MANUAL SORT\\"

def save_file(sp, pdata, header):
    path = sp + header + '.nc'
    with open(path, 'w') as file:
        file.write(pdata)
        file.close()

def files_to_send(jobnum, mach_n_file):
    f_in_dir_list = glob.glob(mach_n_file + jobnum + '*')
    f2s_list = []
    if len(f_in_dir_list) > 0:
        for e in f_in_dir_list:
            if jobnum in e:
                f2s_list.append(e)
    else:
        return None    
    return f2s_list

def loadData(cmd):
    spaces = {3:"       ", 4:"      ", 5:"     ", 6:"    ", 7:"   ", 8:"  ", 9: " "}
    cmd = cmd + spaces.get(len(cmd))    
    MESSAGE = "%CLOD    " + cmd + "\r\n%"    
    BYT = MESSAGE.encode()
    s.send(BYT)
    
def readReply():
    reply = ""
    count = 0
    while count < 2:
        data = s.recv(BUFFER_SIZE)
        reply = reply + data.decode()
        count = reply.count('%')
    return reply

def testCommand(cmd): #this is used to send misc commands to see the machine response
    loadData(cmd)
    reply = readReply()
    print(reply)

def read_file(path): #can be used with either tool path or program path
    with open(path) as file:
        data = file.read()
    return data
        
def get_operation_state():
    state_dict = {0: 'Not Operating', 1: 'Operating', 2: 'Pause', 3: 'Block Stop'}
    mode_dict = {0: 'MANU', 1: 'MDI', 2: 'MEM', 3: 'EDIT', 4: 'MANU+MDI', 5: 'MEM+EDIT'}
    loadData('MEM')
    reply = readReply()
    replist = reply.split(',')
    op_state = state_dict[int(replist[2])]
    op_mode = mode_dict[int(replist[5])]
    return op_state, op_mode

def get_tools():
    loadData('TOLNI1')
    tdata = readReply()
    start = tdata.find('\r\n')
    end = tdata.find('M99,,,,,,,,')
    end = tdata.find('\r\n',end)        
    return tdata[start+2:end]

def get_atc():
    loadData('ATCTL')
    data = readReply()
    start = data.find('\r\n')
    end = data.find('M51,  0,0, 0,1,0\r\n')
    data = data[start+2:end+18]    
    return data

def get_all_dir():
    MESSAGE = "%CDRQALL           \r\n%"
    BYT = MESSAGE.encode()
    s.send(BYT)
    return readReply().split()

def get_prog_list_at_mach():
    data = get_all_dir()
    start = 2
    end = 3
    for i, elem in enumerate(data):
        if 'KM1111' in elem:
            end = i
    plist = []
    for e in data[start:end]:
        e = e[e.find('O'):]
        plist.append(e)
    return plist

def get_cust_part_from_mac(program):
    loadData(program)
    data = readReply()
    start = data.find('(')
    end = data.find(')')
    m30 = data.find('M30')
    return data[start:m30+3], data[start+1:end]    

def send_Data(data, cmd):
    spaces = {3:"       ", 4:"      ", 5:"     ", 6:"    ", 7:"   ", 8:"  ", 9: " "}
    cmd = cmd + spaces.get(len(cmd))
    header = "%CSAV    " + cmd + "\r\n"    
    footer = "\n\r%"
    MESSAGE = header + data + footer
    BYT = MESSAGE.encode()
    s.send(BYT)    

def sendTool(tooldata, cmd="TOLNI1"):
    send_Data(tooldata, cmd)    
    readReply()

def sendATC(atcdata, cmd="ATCTL"):
    send_Data(atcdata, cmd)
    readReply()
    
def sendProg(pnum, jobnum, pp):
    program_list = files_to_send(jobnum, pp)    
    pnum = edit_pnum(pnum)
    if program_list is not None:
        for path in program_list:            
            temp = pnum[1:]
            pnum = edit_pnum(str(int(temp) + program_list.index(path)))            
            program = read_file(path)
            program = editProg(program, jobnum)            
            send_Data(program, pnum)
            reply = readReply()
            if reply[17:19] != '00':
                return "Error sending program"
            else:            
                return "Sending program files success."
    else:
        return "No program found."
    

def edit_pnum(pnum):      
    if pnum[0] != 'O':
        pnum = 'O' + pnum
    if any(i for i in pnum[1:] if i not in '0123456789'):
        return "Bad program number."
    while len(pnum) < 5:        
        pnum = pnum[0] + '0' + pnum[1:]
    return pnum

def editProg(program, jobnum):
    job_line = program.find("(JOB")
    if job_line > 0: #if (JOB ) is found, replace with current job #
        job_end = program.find(")",job_line)
        output = program[:job_line] + '(JOB ' + jobnum + ')\r\n' + program[job_end+2:]        
    else: #if (JOB ) is not found, add job #
        start = program.find('(')
        index = program.find('(',start+1)
        output = program[:index] + '(JOB ' + jobnum + ')\r\n' + program[index:]
    
    m211_index = output.find("M211")
    if m211_index > 0: #M211 is in the program
        return output
    else:
        m30_index = output.find("M30")
        output = output[:m30_index] + "M211\r\n" + output[m30_index:]    
    
    return output        

def delProg(program):
    header = "%CDEL    "
    footer = "     \r\n%"
    MESSAGE = header + program + footer
    BYT = MESSAGE.encode()
    s.send(BYT)
    return readReply()

def merge_tool_files(tool_data, job_number, tp):
    if files_to_send(job_number, tp) is None:
        return tool_data, "No tool presets found."
    
    fp = files_to_send(job_number, tp)[0]
    new_data = read_file(fp).split('\n')
    tool_data_list = tool_data.split('\n')
    tool_dict = {}
    for tool in tool_data_list:
        tool_dict[tool[:3]] = tool
    tool_data_modified = ""
    #print(tool_data_list[0])
    for e in new_data:
        index = new_data.index(e)        
        start = e.find('T')
        if e[start:start+3] != '':
            tool_dict[e[start:start+3]] = e[start:] + '\r'
        tool_data_list[index] = e[start:] + '\r'
    for each in tool_dict:
        if each != 'M99':
            tool_data_modified += tool_dict[each] + '\n' #adds the newline back to each line of tool list
        else:
            tool_data_modified += tool_dict[each]
    return tool_data_modified, "Data updated."
    
    

def handle_tools_and_atc(ap, job_number, tp):
    #before saving tool data, we need to process previous save?
    tool_data = get_tools()
    save_file(ap,tool_data,'tool_list.nc')    
    tool_data, message = merge_tool_files(tool_data, job_number, tp)    
    atc_data = get_atc()
    save_file(ap,atc_data,'atc_list.nc')
        
    return tool_data, atc_data, message

def process_program_request(job_number, vmc):
    pp, tp, ap = set_paths(vmc)
    connect_to_machine(vmc)
    program_list = get_prog_list_at_mach()
    for program in program_list: #cycles through every program on machine
        full, head = get_cust_part_from_mac(program) #gets program and header from machine        
        newhead = head.replace('?','')        
        sp = find_save_path(newhead[:5]) # sp as save path per each program
        save_file(sp, full, newhead) #need a try: except setup here?? this line saves program
        #delProg(program)
        #This space reserved for delProg(program) 
        #This needs to check whether program was called from setup tracker            
    opstatus, mode = get_operation_state()
    if opstatus != 'Operating':
        send_p_msg = sendProg('1', job_number, pp)
        disconnect_machine()
        if send_p_msg != "Sending program files success.":
            return send_p_msg + ". Contact your presetter."
        return send_p_msg + " Program sent as O0001."
    else:
        disconnect_machine()
        return vmc + ' is currently running. Please try again when it is not running.'

def process_tool_request(job_number, vmc):
    pp, tp, ap = set_paths(vmc)
    connect_to_machine(vmc)
    tool_file, atc_file, message = handle_tools_and_atc(ap, job_number, tp) #tool data updated with preset values, or return with no change if no presets found    
    opstatus, mode = get_operation_state()
    if message != "No tool presets found.":
        if opstatus != 'Operating':            
            sendTool(tool_file)
            sendATC(atc_file)
            disconnect_machine()
            return 'Tool files sent to ' + vmc + '.'
        else:
            disconnect_machine()
            return vmc + ' is currently running. Please try again when it is not running.'            
    else:
        disconnect_machine()
        return 'No tool files to send to ' + vmc + '. Contact your presetter'



vmc = 'VanDamme'
job_number = '93437'
print(process_program_request(job_number, vmc))

