# # # INVOICE CHECKERS
# # # PURPOSE: THIS SCRIPT CHECKS THROUGH VARIOUS DIFFERNET INVOICE TEMPLATES AND LOOKS FOR PURCHASE ORDER NUMBERS, WHICH IN THEN TRIES TO MATCH AGAINST A GOODS RECIEVED LOG (EXCEL FILE). IF SUCCESSFULLY MATCHED 
# # # (INCLUDING SUCCESSFUL PRICE MATCH) THE SCRIPT MARKS THE GOODS ON THE EXCEL FILE AS RECIEVED AND THEN RENAMES THE INVOICE PDF FILE TO BE READY FOR POSTING TO SAGE200.
####################################################################################################################################################################
####################################################################################################################################################################
### STEP BY STEP PROCESS:
### 1 - LIST ALL PDF FILES IN FOLDER, READY TO BE CHECKED
### 2 - CHECKS EACH FILE WITH PYPDF 3, IF PYPDF3 FAILS TO READ THE FILE, PYTESSERACT (OCR) IS USED IN ITS PLACE
### 3 - IF A PO NUMBER IS FOUND (6 DIGIT NUMBER BEGINNING WITH EITHER 3 OR 4), IT IS CHECKED AGAINST THE LOG TO FIND A POTENTIAL MATCH. 
### 4 - IF NO MATCH IS FOUND, THE PROCESS TERMINATES HERE AND RESETS TO POINT 2 WITH THE NEXT PDF FILE IN THE LIST
### 5 - IF A MATCH IS FOUND IN THE GOODS RECIEVED LOG, THEN THE NEXT STEP IS TO MAKE SURE THE PRICE OF THE INVOICE, MATCHES THE TOTAL COST OF THE GOODS RECIEVED 
### 6 - IF THE BOTH TOTAL COSTS MATCH, THEN THE FILE IS MARKED AS PASSED ON THE EXCEL FILE AND THE PDF FILE IS RENAMED WITH ITS SUP CODE AND INVOICE NUMBERS.
### 7 - IF THE TOTAL COSTS DONT MATCH, THEN THE FILE IS RENAMED WITH 'ISSUE' AT THE END OF ITS ORIGINAL NAME, AND PLACED INTO AN ISSUES FOLDER FOR A PERSON CHECK
### 8 - IF AT ANY POINT, NO PO CODE IS FOUND FROM EITHER PYTESSERACT OR PYPDF3, THEN THE FILE IS MARKED AS 'TO CHECK' FOR PERSON VISUAL ASSESSMENT OF THE INVOCIE
####################################################################################################################################################################
####################################################################################################################################################################
# # # HOW THIS SCRIPT WORKS
# # # Tesseract: Only scans the first page
# # # Update 0.7 to 0.8: ISSUE files no longer get renamed, due to conflicting PO codes
####################################################################################################################################################################
####################################################################################################################################################################
####################################################################################################################################################################
####################################################################### MODULES ####################################################################################
####################################################################################################################################################################
from tkinter import *
from tkinter import filedialog, filedialog
import collections, pytesseract, os, PyPDF3, openpyxl, re, sys, shutil
from openpyxl.utils.cell import get_column_letter
from pdf2image import convert_from_path
from PIL import Image
from datetime import date
####################################################################################################################################################################
####################################################################### TKINTER ATTRIBUTES #########################################################################
####################################################################################################################################################################
today = date.today()
root=Tk()
root.title("THE ORGINAL FACTORY SHOP INVOICE CHECKER")
root.geometry("450x300+700+300")
root.configure(bg="grey")
####################################################################################################################################################################
####################################################################### FOLDER LOCATIONS ###########################################################################
####################################################################################################################################################################
pytesseract.pytesseract.tesseract_cmd = r'C:\\Users\\nbolade\\Desktop\\Main Folder\\Core Files\\Tesseract-OCR\\tesseract.exe'  ### TESSERACT EXE
poppler_path = r'C:\\Users\\nbolade\\Desktop\\Main Folder\\Core Files\\poppler-0.68.0_x86\\poppler-0.68.0\\bin'  ###change this manually in both locations
os.chdir('C:\\Users\\nbolade\\Desktop\\Main Folder\\Invoice checking folder')                         ### MAIN DIRECTORY
FileFolderAllList = os.listdir("C:\\Users\\nbolade\\Desktop\\Main Folder\\Invoice checking folder")   ### FOLDER LOCATION OF THE INVOICES TO BE PROCESSED
MainFolder = "C:\\Users\\nbolade\\Desktop\\Main Folder\\"                                             ### FOLDER LOCATION OF WHERE TO SAVE THE EXCEL FILE (REF - )
ToCheckFolder = "C:\\Users\\nbolade\\Desktop\\Main Folder\\To Check\\"                                ### FOLDER LOCATION OF WHERE TO SAVE THE ISSUES + TO CHECK INVOICES
PassedFolder = "C:\\Users\\nbolade\\Desktop\\Main Folder\\Passed\\"                                   ### FOLDER LOCATION OF WHERE TO SAVE THE PASSED INVOICES
NotOnLogFolder = "C:\\Users\\nbolade\\Desktop\\Main Folder\\Not on log\\"                             ### FOLDER LOCATION OF WHERE TO SAVE THE NOT ON LOG INVOICES
InvoicesInFolderList = []                                                                             ### THE LIST OF PDF FILES IN THE FILEFOLDER LIST (THIS WILL BE LOOPED)
for files in FileFolderAllList:
    if files.endswith('.pdf') or files.endswith('.PDF'):
        InvoicesInFolderList.append(files)
totalInvoicesInFolder = len(InvoicesInFolderList)                                                     ### TO GET THE TOTAL NUMBER OF INVOCIES TO BE PROCESSED (FOR GUI)                                                     

####################################################################################################################################################################
####################################################################### DATE TIME ##################################################################################
####################################################################################################################################################################                                
DateToday = today.strftime("%d.%m.%y")
DateToday2 = today.strftime("%d/%m/%y")
####################################################################################################################################################################
####################################################################### FUNCTIONS ##################################################################################
####################################################################################################################################################################    
def browser():
    global filename
    filename = filedialog.askopenfilename(initialdir = "G:\\Main Folder",title = "Select file",filetypes = (("Excel","*xlsx"),("all files","*.*")))
    labelfileopened.configure(text="File Opened: "+filename)
def Execute():

    try:
        ExcelFile = openpyxl.load_workbook(filename)
    except:
        Filesbeingprocessed.configure(text = "ERROR: no file selected")
        return
    ####################################################################################################################################################################
    ####################################################################### COUNTERS RESET #############################################################################
    ####################################################################################################################################################################
    passedcounter = 0
    issuecounter = 0
    Notonlogcounter = 0
    InvoiceBeingProcessed = 0
    ToCheckCounter = 0 
    ####################################################################################################################################################################
    ###################################################### CONFIGURING TKINTER LABELS TO DISPLAY COUNTERS ##############################################################
    ####################################################################################################################################################################   
    labelpassedcounter.configure(text="Amount of passed: "+str(passedcounter))
    labelissuecounter.configure(text="Amount of issues: "+str(issuecounter))
    labelNotonlogcounter.configure(text="Amount not on log: "+str(Notonlogcounter))
    labeltocheckcounter.configure(text="Amount to check: "+str(ToCheckCounter))
    Filesbeingprocessed.configure(text = "")
    root.update()
    ####################################################################################################################################################################
    ###################################################### ASSIGNING VALUES AND DICTS FROM THE GRN EXCEL FILE ##########################################################
    ####################################################################################################################################################################   
    ExcelSheet = ExcelFile.active
    POListInExcel = []
    SupCodeInExcel = []
    PriceCostInExcel = []
    ListofSumOfAllPONumbers = []
    for cell in ExcelSheet['A']:
        SupCodeInExcel.append(cell.value)
    for cell in ExcelSheet['B']:
        POListInExcel.append(cell.value)
    for cell in ExcelSheet['D']:
        PriceCostInExcel.append(cell.value)
    POListInExcelLen = len(POListInExcel)
    for i in range(1,int(POListInExcelLen+1)):
        ListofSumOfAllPONumbers.append(str(i))
    POSupList = dict(zip(POListInExcel, SupCodeInExcel))
    POCostList = collections.defaultdict(list)
    POPassedList = collections.defaultdict(list)
    PODateList = collections.defaultdict(list)
    for key, value in zip(POListInExcel, PriceCostInExcel):
        POCostList[key].append(value)
    for key, value in zip(POListInExcel, ListofSumOfAllPONumbers):
        POPassedList[key].append(value)
    for key, value in zip(POListInExcel, ListofSumOfAllPONumbers):
        PODateList[key].append(value)
    ####################################################################################################################################################################
    ###################################################### LOOP (FOR THE ACTUAL PROCESSING OF THE INVOCIES)# ###########################################################
    ####################################################################################################################################################################   
    for Invoices in InvoicesInFolderList:
        InvoiceBeingProcessed += 1
        Filesbeingprocessed.configure(text = "Processing file %d/%d" % (InvoiceBeingProcessed,totalInvoicesInFolder))
        root.update()
        InvoicesNoForName = []
        NotIssue = 0
        NotOnLog = 0
        PoteNumberList = []
        SUPCode = []
        pdfFile = open(Invoices,'rb')
        try:
            pdfReader = PyPDF3.PdfFileReader(pdfFile)
            NumberOfPages = pdfReader.numPages
            allpages = []
            for pages in range(NumberOfPages):
                page = pdfReader.getPage(pages)
                pageContent = page.extractText()
                for content in pageContent.split():
                    allpages.append(content)         #need to make an all pages for tesseract so Excel price can be searched back into Tesseract
            for PoteNumbers in allpages:
                g = float()
                try:
                    g = int(PoteNumbers)
                except:
                    pass
                if PoteNumbers.startswith('4') and len(PoteNumbers) == 6 and isinstance(g, int) or PoteNumbers.startswith('3') and len(PoteNumbers) == 6 and isinstance(g, int):
                    PoteNumberList.append(PoteNumbers)
        except:
            pass
        if PoteNumberList == []:  #### rather than close, use the tesseract here to build a Pote number list of use (each pdf file will go to tesseract if PDF fails
            allpages = []
            page = convert_from_path(Invoices, 200,poppler_path = r'C:\\Users\\nbolade\\Desktop\\Main Folder\\Core Files\\poppler-0.68.0_x86\\poppler-0.68.0\\bin')
            page[0].save(Invoices[:-4]+'.jpg', 'JPEG')
            x = Image.open(Invoices[:-4]+'.jpg')
            pageContent = pytesseract.image_to_string(x)
            for content in pageContent.split():
                allpages.append(content)         #need to make an all pages for tesseract so Excel price can be searched back into Tesseract
            x.close()
            os.remove(Invoices[:-4]+'.jpg')
            for PoteNumbers in allpages:
                g = float()
                try:
                    g = int(PoteNumbers)
                except:
                    pass
                if PoteNumbers.startswith('4') and len(PoteNumbers) == 6 and isinstance(g, int) or PoteNumbers.startswith('3') and len(PoteNumbers) == 6 and isinstance(g, int):
                    PoteNumberList.append(PoteNumbers)
            if PoteNumberList == []:
                pdfFile.close()
                shutil.move(Invoices,ToCheckFolder+Invoices[:-4]+' CHECK'+'.pdf')
                ToCheckCounter
                ToCheckCounter += 1
                labeltocheckcounter.configure(text="Amount to check: "+str(ToCheckCounter))
                continue
        for numbers in PoteNumberList:
            if numbers in POListInExcel:
                if numbers == DuplicatePONumber:
                    continue
                TotalPoCost = sum(POCostList.get(numbers))
                InvoicesNoForName.append(numbers)
                SUPCode = POSupList.get(numbers)
                NotIssue = 1
                if str(f'{TotalPoCost:.2f}') in allpages or str(f'{TotalPoCost:,.2f}') in allpages or str('£'+f'{TotalPoCost:,.2f}') in allpages or str('£'+f'{TotalPoCost:.2f}') in allpages:  #make sure to switch all pages back to Tesseract
                    for s in PODateList.get(numbers):
                        ExcelSheet['F%d' % int(s)].value = DateToday2
                        ExcelSheet['E%d' % int(s)].value = "Y"
                    DuplicatePONumber = numbers 
                    NotIssue = 2
            else: 
                NotOnLog = 1
        if NotIssue == 2:
            pdfFile.close()
            passedcounter += 1
            labelpassedcounter.configure(text="Amount of passed: "+str(passedcounter))
            try:
                shutil.move(Invoices,PassedFolder+"".join(SUPCode)+" "+str(InvoicesNoForName)[2:-2]+'.pdf')
            except:
                try:
                    shutil.move(Invoices,PassedFolder+"".join(SUPCode)+" "+str(InvoicesNoForName)[2:-2]+' 2.pdf')
                except:
                    shutil.move(Invoices,PassedFolder+"".join(SUPCode)+" "+str(InvoicesNoForName)[2:-2]+' 3.pdf')
        elif NotIssue == 1:
            pdfFile.close()
            issuecounter += 1
            labelissuecounter.configure(text="Amount of issues: "+str(issuecounter))
            try:
                shutil.move(Invoices,ToCheckFolder+Invoices[:-4]+' ISSUE'+'.pdf')
            except:
                try:
                    shutil.move(Invoices,ToCheckFolder+Invoices[:-4]+' ISSUE'+' 2.pdf')
                except:
                    shutil.move(Invoices,ToCheckFolder+Invoices[:-4]+' ISSUE'+' 3.pdf')
        elif NotOnLog == 1:
            pdfFile.close()
            Notonlogcounter += 1
            labelNotonlogcounter.configure(text="Amount not on log: "+str(Notonlogcounter))
            try:
                shutil.move(Invoices,NotOnLogFolder+Invoices)
            except:
                try:
                    shutil.move(Invoices,NotOnLogFolder+Invoices[:-4]+' 1'+'.pdf')
                except:
                    shutil.move(Invoices,NotOnLogFolder+Invoices[:-4]+' 2'+'.pdf')
            continue
    ExcelFile.save(MainFolder+"GRN DOWNLOAD %s COMPLETED.xlsx" % DateToday)
    Filesbeingprocessed.configure(text = "All files proccessed %d/%d" % (totalInvoicesInFolder,totalInvoicesInFolder))
    labelpassedcounter.configure(text="Amount of passed: "+str(passedcounter))
    labelissuecounter.configure(text="Amount of issues: "+str(issuecounter))
    labelNotonlogcounter.configure(text="Amount not on log: "+str(Notonlogcounter))
    labeltocheckcounter.configure(text="Amount to check: "+str(ToCheckCounter))
####################################################################################################################################################################
####################################################################### TKINTER LABELS #############################################################################
####################################################################################################################################################################    
label_file_explorer = Label(root,text = "Original Factory Shop Invoice Checker - By Nick",width = 100, height = 4,fg = "blue",bg = 'grey')
button_explore = Button(root,text = "Browse Files",bg = 'grey',width = 10, height = 1,command = browser)
button_exit = Button(root,text = "Exit",bg = 'grey',width = 10,height = 1,command = sys.exit)
button_execute = Button(root,text = "Execute",bg = 'grey',width = 10,height = 1,command = Execute)
labelfileopened = Label(root,text = "",width = 50, height = 1,fg = "blue",bg = 'grey')
labelpassedcounter = Label(root,text = "",width = 20,fg = "blue",bg = 'grey')
labelissuecounter = Label(root,text = "",width = 20,fg = "blue",bg = 'grey')
labeltocheckcounter = Label(root,text = "",width = 20,fg = "blue",bg = 'grey')
labelNotonlogcounter = Label(root,text = "",width = 20,fg = "blue",bg = 'grey')
Filesbeingprocessed = Label(root,text = "",bg = 'grey')
####################################################################################################################################################################
####################################################################### TKINTER PACKS #############################################################################
####################################################################################################################################################################    
def allpacks():
    label_file_explorer.pack()
    button_explore.pack()
    button_execute.pack()
    button_exit.pack()
    labelfileopened.pack()
    Filesbeingprocessed.pack()
    labelpassedcounter.pack()
    labelissuecounter.pack()
    labelNotonlogcounter.pack()
    labeltocheckcounter.pack()
allpacks()
root.mainloop()
