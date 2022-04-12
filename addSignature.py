import os, openpyxl
from win32com import client
from openpyxl.drawing.image import Image
from pathlib import Path

DEBUG = False # DEBUG
INFO = True # INFO

# Header
print("")
print("==================== Podepisovač docházek ====================")
print("")

# Inputs
year = input("Enter the year: ")
month = input("Enter the month: ")
print("")
print("--------------------------------------------------------------")
print("")

# Variables
prefix = (year + "-" + month + "-")
dir_in = os.path.abspath('in\\')
dir_out = os.path.abspath('out\\')
img_signature = Image('signature-alpha.png')
n = 0

# Create subfolders if not exist
Path(dir_in).mkdir(parents=True, exist_ok=True)
Path(dir_out).mkdir(parents=True, exist_ok=True)

if DEBUG: print(prefix, dir_in, dir_out, n) # DEBUG

# Loop to process every xlsx file in input directory
for file in os.listdir(dir_in):
    fname = os.fsdecode(file)
    if fname.endswith(".xlsx"): 
        inFile = (dir_in + "\\" + fname)
        outFile = (dir_out + "\\" + fname)
        
        # Open file, insert signature, save to output folder and close
        if INFO: print("Read: " + inFile) # INFO
        wb = openpyxl.load_workbook(filename = inFile)
        ws = wb.active
        user = ws['C4'].value
        if DEBUG: print("User: " + user) # DEBUG
        if DEBUG: print("Add signature image") # DEBUG
        ws.add_image(img_signature, 'J46')
        if INFO: print("Write: " + outFile) # INFO
        wb.save(filename = outFile)
        wb.close()
        
        # Load output excel file and save to PDF
        pdfFile = (dir_out + "\\" + prefix + user + ".pdf")
        if INFO: print("Write: " + pdfFile) # INFO
        excel = client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(outFile)
        ws = wb.Worksheets[0]
        ws.ExportAsFixedFormat(0, pdfFile)
        wb.Close(False)        
      
        # Delete exels file
        if INFO: print("Delete: " + inFile) # INFO
        if os.path.exists(inFile): os.remove(inFile)
        if INFO: print("Delete: " + outFile) # INFO
        if os.path.exists(outFile): os.remove(outFile)
        
        n = n + 1 # Increment loop
        if INFO: print(user + " done.") # INFO
        print("")
        print("--------------------------------------------------------------")
        print("")

# Footer
print('Completed.\n%d attendance sheets signed.' % n)
print("")
print("==============================================================")
print("")

os.system("pause") # press any key