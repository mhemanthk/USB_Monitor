import wmi
import os
from datetime import date, datetime
from os.path import exists

#Define Drive
USBDrive = 'D:'

#Program Started
now = datetime.now()
dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
print('Program_started @ ' + dt_string)

#Define Headers
headers = "Time, USB Action , DeviceName, DeviceID, SystemName, Files"

#Define the file path here
dirpath = str('C:\\USBMon\\')
print('dirpath = ', dirpath)

#WMI call
device_action_wql = "SELECT * FROM __InstanceOperationEvent WITHIN 4 WHERE TargetInstance ISA \'Win32_USBHub\'"
c = wmi.WMI()
device_action = c.watch_for(raw_wql=device_action_wql)


#USB Aud python program log
logDirPath = 'C:\\USBAud\\'
if not os.path.exists(logDirPath):
    os.makedirs(logDirPath)
with open('C:\\USBAud\\USBAudPythonPgmLog.txt', "a") as file:
    file.write('Program_started @ ' + dt_string)
    file.write("\n")



#Create Directory
def createDir():
    if not os.path.exists(dirpath):
        os.makedirs(dirpath)

def newtime():
    now = datetime.now()
    dt_filename = now.strftime("%d-%m-%Y")
    print('new Date :',dt_filename)
    return  dt_filename



#Create file everyday
def createfile(dt_filename):
    filename = dt_filename + '.csv'
    print('filename = ',filename)
    fullpath = dirpath + filename
    print('fullpath = ', fullpath)
    print(fullpath)
    if exists(fullpath):
        print('file_exists')
    else:
        with open(fullpath, "a") as file:
            file.write(headers)
            file.write("\n")
    return fullpath


#Log the USB events
def logAction(fullpath, output):
    with open(fullpath, "a") as file:
        file.write(output)
        file.write("\n")


def read_usb():
        if os.path.isdir(USBDrive):
            contents = os.listdir(USBDrive)
            if len(contents) > 1:
                return contents
            else:
                return 0

print('-'*72)

def keeprunning():
    while True:
        try:
            usb_action = device_action()
            #print(usb_action)
            DevNamew2f = usb_action.Name
            type(DevNamew2f)
            print("DeviceName :", usb_action.Name)
            DevIDw2f = usb_action.DeviceID
            type(DevIDw2f)
            print("DeviceID :", usb_action.DeviceID)
            SysNamew2f = usb_action.SystemName
            type(SysNamew2f)
            print("SystemName :", usb_action.SystemName)
            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
            #print("USB action @ =", dt_string)
            files = read_usb()
            #print(files)
            if files == None:
                fileStatw2F = "No File @ USB Ejected"
                type(fileStatw2F)
                Actionw2f = "USB Ejected "
                timestampw2f = dt_string
                type(Actionw2f)
                print("USB Ejected @ ", dt_string)
            else:
                Actionw2f = "USB Inserted "
                timestampw2f = dt_string
                type(Actionw2f)
                print("USB Inserted @ ", dt_string)
                if files == 0:
                    fileStatw2F = "No files"
                    type(fileStatw2F)
                    print('No files')
                else:
                    fileStatw2F = "Files are : " + str(files)
                    type(fileStatw2F)
                    print('Files are :', files)
            outputw2f = str(timestampw2f) + ', '+ str(Actionw2f) + ', '+ str(DevNamew2f) + ', '+ str(DevIDw2f) + ', ' + str(SysNamew2f) + ', '+  str(fileStatw2F)
            print('*' * 72)
            createDir()
            dt_filename = newtime()
            fullpath = createfile(dt_filename)
            logAction(fullpath, outputw2f)
            print('*' * 72)
        except:
            with open('C:\\USBAud\\USBAudPythonPgmLog.txt', "a") as file:
                file.write('Exception @ ' + dt_string)
                file.write("\n")


keeprunning()










