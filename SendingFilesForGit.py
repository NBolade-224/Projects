###### EMAIL SENDER

### need to set automatic creation of ready to send folder and sent folder. 
from tkinter import *
from tkinter import filedialog
from pathlib import Path
import win32com.client as win32
import collections, pytesseract, os, PyPDF3, openpyxl, re, sys, shutil
from openpyxl.utils.cell import get_column_letter
from pdf2image import convert_from_path
from PIL import Image
import time
pytesseract.pytesseract.tesseract_cmd = r'C:\\Users\\nickb\\Desktop\\Python\\Tesseract-OCR\\tesseract.exe'  ####  this will be compiled? 
root=Tk()
root.title("Sending Emails")
root.geometry("450x450+700+300")
root.configure(bg="grey")
FilesinFolderlist = []
FilenametoSend = []
############################
############################
############################
############################
###########################
############################
############################
TotalFilesinFolder = 0
FilenameentCOunter = 0
FilesRemaned = 0
filestosend = 0
global Safety
Safety = 'On'
poppler_path = r'C:\\Users\\nickb\Desktop\\Python\\poppler-0.68.0_x86\\poppler-0.68.0\\bin'  ###  his will be compiled? 
def browser():
    labelforERRORs.configure(text='')
    global filename
    global ExcelFile
    global ExcelSheet
    global EmailAddressList
    global ReferenceNumberList
    global NamesList
    global RefnumbertoEmail
    global EmailtoNameDict
    global RefinNumberList
    global POtoRefinNumberlist
    EmailAddressList = []
    ReferenceNumberList = []
    NamesList = []
    RefinNumberList = []
    try:
        filename = filedialog.askopenfilename(initialdir = "C:\\Users\\nickb\\Desktop",title = "Select file",filetypes = (("Excel","*xlsx"),("all files","*.*")))
        ExcelFile = openpyxl.load_workbook(filename)
        ExcelSheet = ExcelFile.active
        labelfileopened.configure(text="File Opened: "+filename)
    except:
        labelforERRORs.configure(text='Please select a valid excel file')
        return
    for cell in ExcelSheet['G']:
        EmailAddressList.append(cell.value)
    for cell in ExcelSheet['C']:
        ReferenceNumberList.append(cell.value)
    for cell in ExcelSheet['F']:
        RefinNumberList.append(cell.value)    
    for cell in ExcelSheet['B']:
        NamesList.append(cell.value)
    RefnumbertoEmail = dict(zip(ReferenceNumberList, EmailAddressList))
    EmailtoNameDict = dict(zip(EmailAddressList, NamesList))
    POtoRefinNumberlist = dict(zip(ReferenceNumberList, RefinNumberList))
    root.update()
def browser2():
    global FilesinFolderlist
    global TotalFilesinFolder
    global Fildesfolder
    FilesinFolderlist = []
    Fildesfolder = filedialog.askdirectory(initialdir = "C:\\Users\\nickb\\Desktop",title = "Select File Folder")
    os.chdir(Fildesfolder) 
    FilenameInFolder = os.listdir(Fildesfolder)
    for files in FilenameInFolder:
        if files.endswith('.pdf') or files.endswith('.PDF'):
            FilesinFolderlist.append(files)
    TotalFilesinFolder = len(FilesinFolderlist)
    labetotalP45tosend.configure(text="Files to be processed: "+str(TotalFilesinFolder))
    root.update()
def Sort():
    ListofFilenametoEmailAddress = []
    labelforERRORs.configure(text="")
    ReferenceNumber = ""
    EmailAddress = ""
    Name = ""
    labetotalP45tosend.configure(text="")
    global FilesinFolderlist
    global FilesRemaned
    global FilenametoSend
    global Safety
    Safety = 'On'
    FilenametoSend = []
    FilesRemaned = 0
    Error = False
    for Filename in FilesinFolderlist:
        ReferenceNumber = ""
        EmailAddress = ""
        Name = ""
        RefCode = ""
        newFileName = ''
        page = convert_from_path(Filename, 350,poppler_path = r'C:\\Users\\nickb\\Desktop\\Python\\poppler-0.68.0_x86\\poppler-0.68.0\\bin')  ## will be compiled into exe
        page[0].save(Filename[:-4]+'.jpg', 'JPEG')
        x = Image.open(Filename[:-4]+'.jpg')
        pageContent = pytesseract.image_to_string(x)
        for content in pageContent.split():
            if len(content) == 6 and content.startswith("4"): ## and is int
                ReferenceNumber = content
        if ReferenceNumber == "":
            labelforERRORs.configure(text="ERROR: Failed to find payroll number in file: "+Filename)
            Error = True
            x.close()
            os.remove(Filename[:-4]+'.jpg')
            break
        EmailAddress = RefnumbertoEmail.get(ReferenceNumber)
        Name = EmailtoNameDict.get(EmailAddress)
        RefCode =  POtoRefinNumberlist.get(ReferenceNumber)
        if RefCode in pageContent.split():
            pass
        else:
            labelforERRORs.configure(text="ERROR: National Insurance Number does not match report: "+Filename)
            x.close()
            os.remove(Filename[:-4]+'.jpg')
            Error = True
            break
        x.close()
        os.remove(Filename[:-4]+'.jpg')
        PDFFile = PyPDF3.PdfFileReader(Fildesfolder+'/'+Filename)
        NumberOfPages = PDFFile.numPages
        Output_PDFFile = PyPDF3.PdfFileWriter()
        for i in range(NumberOfPages):
            Output_PDFFile.addPage(PDFFile.getPage(i))
        Output_PDFFile.encrypt(RefCode)
        Output_PDFFile.write(open(Fildesfolder+'/FilenameRenamed\\'+ReferenceNumber+" "+Name+".pdf", 'wb'))
        FilesRemaned += 1
        newFileName = ReferenceNumber+" "+Name
        filesconvertedcounter.configure(text = "Amount of files coverted: "+str(FilesRemaned))
        FilenametoSend.append(Fildesfolder+'/FilenameRenamed/'+ReferenceNumber+" "+Name+".pdf")
        ListofFilenametoEmailAddress.append(newFileName+'    ---->    '+EmailAddress)
        newFileName = ''
        RefCode = ""
        ReferenceNumber = ""
        EmailAddress = ""
        Name = ""
        labetotalP45tosend.configure(text='File being Processed: '+str(FilesRemaned)+'/'+str(TotalFilesinFolder))
        root.update()
    if Error == False:
        labetotalP45tosend.configure(text="Total converted is "+str(len(FilenametoSend)))
        window = Toplevel(root)
        window.title('List of Filename and Address for them to be sent to')
        window.geometry("450x300+700+400")
        ListofP45toEmail = Listbox(window, bg="darkgrey",width=50, height=50, selectmode='single')
        ListofP45toEmail.pack()
        ListofP45toEmail.insert(END,'Total Filename to be sent '+str(TotalFilesinFolder))
        ListofP45toEmail.insert(END,'')
        for x in ListofFilenametoEmailAddress:
            ListofP45toEmail.insert(END,x)
        Safety = 'Off'
        root.update()
    else:
        pass
def Send():
    global Safety
    if Safety == 'Off':
        ReferenceNumber = ""
        EmailAddress = ""
        Name = ""
        global FilesRemaned
        global FilenametoSend
        global FilenameentCOunter
        for Filenametos in FilenametoSend:
            ReferenceNumber = ""
            EmailAddress = ""
            Name = ""
            RefCode = ""
            file = Path(Fildesfolder+Filenametos).stem           #### CHANGE FOLDER NAME HERE 
            ReferenceNumber = file[:6]
            FileName = file[7:]
            EmailAddress = RefnumbertoEmail.get(ReferenceNumber)
            Name = EmailtoNameDict.get(EmailAddress)
            RefCode =  POtoRefinNumberlist.get(ReferenceNumber)
            if FileName == Name:
                pass
            else:
                break
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = EmailAddress
            mail.Subject = 'Test'
            mail.Body = 'Testing Email....\n pw is '+str(RefCode)
            mail.Attachments.Add(Filenametos)
            mail.Send()
            FilenameentCOunter += 1
            labelsentcounter.configure(text = "Amount of files sent: %d/%d" % (FilenameentCOunter,TotalFilesinFolder))
            RefCode = ""
            ReferenceNumber = ""
            EmailAddress = ""
            Name = ""
            time.sleep(3)
            root.update()
        labelsentcounter.configure(text = "All emails sent %d/%d" % (FilenameentCOunter,TotalFilesinFolder))
        root.update()
    else:
        labelforERRORs.configure(text='Please sort the Filename before attempting to send')
label_file_explorer = Label(root,text = "By Nick",width = 100, height = 4,fg = "blue",bg = 'grey')
button_explore = Button(root,text = "Select Report",bg = 'grey',width = 30, height = 2,command = browser)
button_explore2 = Button(root,text = "Select Folder",bg = 'grey',width = 30, height = 2,command = browser2)
button_exit = Button(root,text = "Exit",bg = 'grey',width = 30,height = 2,command = sys.exit)
button_sort = Button(root,text = "Sort",bg = 'grey',width = 30,height = 2,command = Sort)
button_send = Button(root,text = "Send",bg = 'grey',width = 30,height = 2,command = Send)
labelfileopened = Label(root,text = "",width = 50, height = 2,fg = "blue",bg = 'grey')
labetotalP45tosend = Label(root,text = "",width = 50, height = 2,fg = "blue",bg = 'grey')
labelforERRORs = Label(root,text = "",width = 50, height = 2,fg = "blue",bg = 'grey')
labelsentcounter = Label(root,text = "",width = 20,fg = "blue",bg = 'grey')
filesconvertedcounter = Label(root,text = "",width = 20,fg = "blue",bg = 'grey')
Filesbeingprocessed = Label(root,text = "",bg = 'grey')
def allpacks():
    label_file_explorer.pack()
    button_explore.pack()
    button_explore2.pack()
    button_sort.pack()
    button_exit.pack()
    button_send.pack()
    labelforERRORs.pack()
    labelfileopened.pack()
    labetotalP45tosend.pack()
    labelsentcounter.pack()
allpacks()
root.mainloop()


