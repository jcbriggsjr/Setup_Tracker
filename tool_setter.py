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

strikepath = 'Striking\\'
probepath = 'Probing\\'

BUFFER_SIZE = 1024

def connect_to_machine(machine):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    TCP_IP = machineip[machine]
    TCP_PORT = 10000    
    s.connect((TCP_IP, TCP_PORT))         
    return s

def disconnect_machine(s):
    s.close()


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
    search_dir = 'G:\\3 - Production Departments\\4 - Grinding\\9 - VMCs\\4 - Programs & Software\\2 - Programs & MC Files'    
    dir_list = os.listdir(search_dir)
    folder = [s for s in dir_list if customer.lower() in s.lower()]
    if len(folder) > 0:
        return search_dir + '\\' + folder[0] + '\\'
    return search_dir + "\\MANUAL SORT\\"

def save_file(sp, pdata, header):
    path = sp + header + '.nc'
    with open(path, 'w') as file:
        file.write(pdata)
        file.close()

def files_to_send(jobnum, filepath):
    f_in_dir_list = glob.glob(filepath + jobnum + '*')
    f2s_list = []
    if len(f_in_dir_list) > 0:
        for e in f_in_dir_list:
            if jobnum.upper() in e.upper():
                f2s_list.append(e)
    else:
        return None    
    return f2s_list

def loadData(cmd, s):
    spaces = {3:"       ", 4:"      ", 5:"     ", 6:"    ", 7:"   ", 8:"  ", 9: " "}
    cmd = cmd + spaces.get(len(cmd))    
    MESSAGE = "%CLOD    " + cmd + "\r\n%"    
    BYT = MESSAGE.encode()
    s.send(BYT)
    
def readReply(s):
    reply = ""
    count = 0
    while count < 2:
        data = s.recv(BUFFER_SIZE)
        reply = reply + data.decode()
        count = reply.count('%')
    return reply

def testCommand(cmd, s): #this is used to send misc commands to see the machine response
    loadData(cmd)
    reply = readReply(s)
    print(reply)

def read_file(path): #can be used with either tool path or program path
    with open(path) as file:
        data = file.read()
    return data
        
def get_operation_state(s):
    state_dict = {0: 'Not Operating', 1: 'Operating', 2: 'Pause', 3: 'Block Stop'}
    mode_dict = {0: 'MANU', 1: 'MDI', 2: 'MEM', 3: 'EDIT', 4: 'MANU+MDI', 5: 'MEM+EDIT'}    
    loadData('MEM', s)
    reply = readReply(s)
    replist = reply.split(',')
    op_state = state_dict[int(replist[2])]
    op_mode = mode_dict[int(replist[5])]
    return op_state, op_mode

def get_tools(s):
    loadData('TOLNI1', s)
    tdata = readReply(s)
    start = tdata.find('\r\n')
    end = tdata.find('M99,,,,,,,,')
    end = tdata.find('\r\n',end)        
    return tdata[start+2:end]

def get_atc(s):
    loadData('ATCTL', s)
    data = readReply(s)
    start = data.find('\r\n')
    end = data.find('M51,  0,0, 0,1,0\r\n')
    data = data[start+2:end+18]    
    return data

def get_all_dir(s):
    MESSAGE = "%CDRQALL           \r\n%"
    BYT = MESSAGE.encode()
    s.send(BYT)
    reply = readReply(s).split()
    return reply

def get_prog_list_at_mach(s):
    data = get_all_dir(s)
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

def get_cust_part_from_mac(program, s):
    loadData(program, s)
    data = readReply(s)
    start = data.find('(')
    end = data.find(')')
    m30 = data.find('M30')
    return data[start:m30+3], data[start+1:end]    

def send_Data(data, cmd, s):
    spaces = {3:"       ", 4:"      ", 5:"     ", 6:"    ", 7:"   ", 8:"  ", 9: " "}
    cmd = cmd + spaces.get(len(cmd))
    header = "%CSAV    " + cmd + "\r\n"    
    footer = "\n\r%"
    MESSAGE = header + data + footer
    BYT = MESSAGE.encode()
    s.send(BYT)    
    
def sendTool(tooldata, s, cmd="TOLNI1"):
    send_Data(tooldata, cmd, s)    
    readReply(s)

def sendATC(atcdata, s, cmd="ATCTL"):
    send_Data(atcdata, cmd, s)
    readReply(s)

# send_strikeprobe requires program file as a string, program number to send
# and s as the network connection
def send_StrikeProbe(progdata, prognumber, s):
    # send_Data handles sending the actual data over. 
    send_Data(progdata,prognumber,s)
    # check reply to see if we were successful in sending program
    reply = readReply(s)
    if reply[17:19] != '00':
        return "Error sending program"
    else:            
        return "Sending program files success."


    
def sendProg(jobnum, pp, s):
    program_list = files_to_send(jobnum, pp)    
    pnum = 1001
    if program_list is not None:
        for path in program_list:            
            pnum_str = 'O' + str(pnum)
            program = read_file(path)
            program = editProg(program, jobnum)
            send_Data(program, pnum_str, s)
            reply = readReply(s)
            pnum += 1
            if reply[17:19] != '00':
                return "Error sending program"
            
        return "Sending program files success."
    else:
        return "No program found."

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
    m213_index = output.find("M213")
    if m211_index > 0: #M211 is in the program
        if m213_index > 0: #M213 is also in the program
            return output
        else:
            m30_index = output.find("M30")
            output = output[:m30_index] + "M213\r\n" + output[m30_index:]    
    else:
        m30_index = output.find("M30")
        output = output[:m30_index] + "M211\r\nM213\r\n" + output[m30_index:]    
    
    return output        

def delProg(program, s):
    header = "%CDEL    "
    footer = "     \r\n%"
    MESSAGE = header + program + footer
    BYT = MESSAGE.encode()
    s.send(BYT)
    return readReply(s)

def merge_tool_files(tool_data, job_number, tp):    
    if files_to_send(job_number, tp) is None or len(job_number) <5:
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
        stop = e.find(',',start)     
        dl = len(e[start:stop]) #data length
        if dl != 3 or e[start+1] not in '0123456789' or e[start+2] not in '0123456789':
            if dl !=0:        
                return tool_data, "Bad tool data. Contact programmer."
        if e[start:start+3] != '':
            tool_dict[e[start:start+3]] = e[start:] + '\r'
        tool_data_list[index] = e[start:] + '\r'
    for each in tool_dict:
        if each != 'M99':
            tool_data_modified += tool_dict[each] + '\n' #adds the newline back to each line of tool list
        else:
            tool_data_modified += tool_dict[each]
    return tool_data_modified, "Tool data sent."
    
    

def handle_tools_and_atc(ap, job_number, tp, s):
    #before saving tool data, we need to process previous save?
    tool_data = get_tools(s)
    save_file(ap,tool_data,'tool_list.nc')    
    tool_data, message = merge_tool_files(tool_data, job_number, tp)    
    atc_data = get_atc(s)
    save_file(ap,atc_data,'atc_list.nc')
        
    return tool_data, atc_data, message

def process_program_request(job_number, vmc):
    pp, tp, ap = set_paths(vmc)
    s = connect_to_machine(vmc)
    program_list = get_prog_list_at_mach(s)
    for program in program_list: #cycles through every program on machine
        full, head = get_cust_part_from_mac(program, s) #gets program and header from machine        
        newhead = head.replace('?','')
        sp = find_save_path(newhead[:5]) # sp as save path per each program
        try:
            save_file(sp, full, newhead) #need a try: except setup here?? this line saves program
        except:
            print("program " + newhead[9:15] + " empty")
        delProg(program, s)
        #This space reserved for delProg(program) 
        
    opstatus, mode = get_operation_state(s)
    # copy program to a pallet 2 version
    modify_prog_pallet2(pp, job_number)
    
    if opstatus != 'Operating':
        send_p_msg = sendProg(job_number, pp, s)
        disconnect_machine(s)
        if send_p_msg != "Sending program files success.":
            return send_p_msg + ". Contact your presetter."
        return send_p_msg + " Programs sent."
    else:
        disconnect_machine(s)
        return vmc + ' is currently running. Please try again when it is not running.'

def process_tool_request(job_number, vmc):
    pp, tp, ap = set_paths(vmc)        
    s = connect_to_machine(vmc)
    tool_file, atc_file, message = handle_tools_and_atc(ap, job_number, tp, s) #tool data updated with preset values, or return with no change if no presets found    
    opstatus, mode = get_operation_state(s)
    if message != "No tool presets found.":
        if opstatus != 'Operating':            
            sendTool(tool_file, s)
            sendATC(atc_file, s)
            disconnect_machine(s)
            return message
        else:
            disconnect_machine(s)
            return vmc + ' is currently running. Please try again when it is not running.'            
    else:
        disconnect_machine(s)
        return 'No tool files to send to ' + vmc + '. Contact your presetter'
    
def process_strikeprobe_request(job_number, vmc, prog_type):
    
    # set path to desired files
    # initialize program number
    if prog_type == 'Striking':
        selectedpath = path + vmc + '\\Striking\\'
        program_number = 2001
    else:
        selectedpath = path + vmc + '\\Probing\\' 
        program_number = 3001
    
    s = connect_to_machine(vmc)
    # get list of files that match job number
    files_list = files_to_send(job_number,selectedpath)    
    opstatus, mode = get_operation_state(s)
    # verify machine is not operating
    if opstatus != 'Operating':
        for filepath in files_list:
            # send each file found in list
            with open(filepath,'r') as file:
                program_data = file.read()
            # cast program number as string and add "O"
            pnum_str = 'O' + str(program_number)
            send_msg = send_StrikeProbe(program_data,pnum_str, s)
            if send_msg != "Sending program files success.":
                return send_msg + ". contact your presetter."
            program_number += 1
        return prog_type + " programs sent successfully."
        disconnect_machine(s)
    else:
        disconnect_machine(s)
        return vmc + ' is currently running. Please try again when it is not running.'
    
def modify_prog_pallet2(pp, job_number):
    # get path to program pallet 1
    program_path = files_to_send(job_number, pp)
    # read pallet 1 program
    with open(program_path[0], 'r') as file:
        pallet2_program = file.read()
    # modify pallet 1 program to save back as pallet 2 program
    # changing which WCS is used
    pallet2_program = pallet2_program.replace("G54.1 P1", "G54.1 P4")
    pallet2_program = pallet2_program.replace("G54.1 P2", "G54.1 P5")
    pallet2_program = pallet2_program.replace("G54.1 P3", "G54.1 P6")
    
    # create program pallet 2 path
    program_path2 = pp + job_number + '_pallet2.nc'
    # save program pallet 2 in same directory as program pallet 1
    with open(program_path2, 'w') as file2:
        file2.write(pallet2_program)
    
    
if __name__ == "__main__":
    pp,tp,ap = set_paths('VanDamme')
    modify_prog_pallet2(pp,'12345')