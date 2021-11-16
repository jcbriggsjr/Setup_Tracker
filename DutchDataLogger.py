from time import sleep
import telnetlib
import sys


destFile = "DutchLog.csv"
getDataCommand = "%CLOD    PRD3      \r\n".encode("ascii")
confirmCommand = "%\r\n".encode("ascii")
getExerCommand = "%CIOCREF EXER10    \r\n".encode("ascii")#Check if button signaling a setup is depressed
delimiter = "%".encode("ascii")
HOST = "192.168.84.217"
PORT = 10000
try:
    with open(destFile, "r") as file:
        mostRecentDateTime = int(file.readlines()[-1].split(",")[0])
except:
    mostRecentDateTime = 0

while True:
    try:
        connection = telnetlib.Telnet(HOST, PORT)
        
        #Check if button signaling setup is on
        connection.write(getExerCommand)
        connection.write(confirmCommand)
        connection.read_until(delimiter, 1)
        data = str(connection.read_until(delimiter, 1), "UTF-8")#newest will be at end of list
        setupButtonState = data.split("\r\n")[1]
        
        connection.write(getDataCommand)
        connection.write(confirmCommand)
        connection.read_until(delimiter, 1)
        data = str(connection.read_until(delimiter, 1), "UTF-8")#newest will be at end of list
        data = data.split("\r\n")[2:-2]#Strip empty lines
        data.reverse()#Newest now first
        data = [line.split(",") for line in data]
        newData = []#We are pulling in the whole file each time
        #so most of the data has already been aquired.
        for line in data:
            dateTime = line[1]
            status = line[2]
            programError = line[4]
            if int(dateTime) > int(mostRecentDateTime):
                firstLine, secondLine, pcPerCycle = "None", "None", 1
                getProgramCommand = "%CLOD    O".encode("ascii")+programError.encode("ascii")+"     \r\n".encode("ascii")
                connection.write(getProgramCommand)
                connection.write(confirmCommand)
                connection.read_until(delimiter, 1)
                program = str(connection.read_until(delimiter, 1), "UTF-8")
                #print(program)
                if len(program) > 24:
                    firstLine, secondLine = program.split("\r\n")[1:3]
                    pcPerCycle += ("G55" in program) + ("G56" in program)
                newline = dateTime+","+status+","+programError+',"'+firstLine+'",'+str(pcPerCycle)+","+setupButtonState+',"'+secondLine+'"\n'
                print(newline)
                newData.append(newline)
                
        print(data[-1][1]+": "+str(len(newData))+ " New lines.")
        mostRecentDateTime = data[-1][1]
        #Comment out the following couple lines while testing so we don't write potentially erronous data.
        with open(destFile, "a") as file:
            file.writelines(newData)
    except Exception as e:
        a, b, c = sys.exc_info()
        print(e, a, b, c)
    try:
        connection.close()
    except:
        print("No Connection")
    sleep(30)
