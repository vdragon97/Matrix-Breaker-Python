from PIL import Image  
import sys
import os
import requests
import pandas as pd
import time
import re
import csv

def prepare_image(img):
    if 'L' != img.mode:
        img = img.convert('L')
    return img

def remove_noise(img, pass_factor):
    for column in range(img.size[0]):
        for line in range(img.size[1]):
            value = remove_noise_by_pixel(img, column, line, pass_factor)
            img.putpixel((column, line), value)
    return img

def remove_noise_by_pixel(img, column, line, pass_factor):
    if img.getpixel((column, line)) < pass_factor:
        return (0)
    return (255)

def call_api(img):
    rurl = "http://api.ocrapiservice.com/1.0/rest/ocr"
    files 	= {	"image"		:("temp.jpg", open(img, "rb"), "image/jpeg")}
    data	= {	"language"	:"en",
                "apikey"	:"MKb65Y2LQW"}
    response = requests.post("http://api.ocrapiservice.com//1.0/rest/ocr", files=files, data=data)
    return response.text
    
def csv_writer(data, file_path):
    csv_file = open(file_path, 'w', newline = '')
    csv_writer = csv.writer(csv_file)
    csv_writer.writerows(data)
    csv_file.close()
    
if __name__=="__main__":
    print ("Step 0: Opening original file matrix JPG")
    path = ".\image"
    input_image = os.path.join(path, sys.argv[1] + '.jpg')
    temp_image = os.path.join(path, sys.argv[1] + '_temp.jpg')
    pass_factor = 100
        
    print ("Step 1: Greyscale converting file matrix JPG")
    img = Image.open(input_image)
    img = prepare_image(img)
    img = remove_noise(img, pass_factor)
    img.save(temp_image)
    
    print ("Step 2: Call API")    
    output = call_api(temp_image)
    output = re.sub('[^A-Z0-9\n ]+', '', output)
    print (output)
    
    
    print ("Step 3: Export to CSV")
    csv_path = os.path.join(path, sys.argv[1] + '.csv')
    output = re.sub('[^A-Z0-9\n]+', '', output)
    content = output.split("\n")
    csv_writer(content, csv_path)
    
    print ("Step 4: Save as XLSX")
    read_file = pd.read_csv (csv_path, names=['A','B','C','D','E','F','G'])
    xlsx_path = os.path.join(path, sys.argv[1] + '.xlsx')
    read_file.index += 1 
    read_file.to_excel (xlsx_path, index = True, header=True)
    
    print ("Complete!")
    
    """
    http://ocrapiservice.com/myaccount/dashboard/
    http://ocrapiservice.com/documentation/testdrive/
    http://ocrapiservice.com/documentation/
    """