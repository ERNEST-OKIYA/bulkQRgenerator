from openpyxl import load_workbook
import qrcode

#Load the excel file
wb = load_workbook("02.xlsx")

#Add the sheet name
sheet_info = wb['juja']

#qr_gen function is responsible for generating QR codes with MEMCard format
def qr_gen(Filename = None, Name = None, Phonenumber = None, RM = None) :
    qFilename   = "FILENAME: " + str(Filename)
    qName    = "NAME: " + str(Name)
    
    img = qrcode.make(f"QR Code Number:{Filename}; Full Name:{Name}; Phone Number:{Phonenumber}; Relationship Manager:{RM};")
    type(img)
    img.save(f"qrfiles/{Filename}.png")

#Bulk processing
for row in range(2, 1697) :
    str_row = str(row)
    qr_gen(
        sheet_info["A" + str_row].value,
        sheet_info["B" + str_row].value,
        sheet_info["C" + str_row].value,
        sheet_info["D" + str_row].value,
    )
