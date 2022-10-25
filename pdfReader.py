from PIL import Image
from pytesseract import pytesseract
import xlsxwriter
import os


workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()

#Define path to tessaract.exe
path_to_tesseract = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

#Define path to image
path_to_images = r'C:/Users/Dell/Desktop/test/'

#Point tessaract_cmd to tessaract.exe
pytesseract.tesseract_cmd = path_to_tesseract

row = 0
for root, dirs, file_names in os.walk(path_to_images):
    #Iterate over each file_name in the folder
    for file_name in file_names:
        #Open image with PIL
        img = Image.open(path_to_images + file_name)

        #Extract text from image
        text = pytesseract.image_to_string(img)
        
        result = text[text.find('To:')+3:text.find('Subj:')]
        print(result)
        
        words = result.split('\n')
        
        for col_num, data in enumerate(words):
            worksheet.write(row, col_num, data)
            
        row=row+1


workbook.close()