import sys
import pdfquery
import pandas as pd
import xlsxwriter

def getGenInfo(file):
    genInfoLabel  = [
    "Agency E-mail Address:", 
    "Website Address:", 
    "Legal Agency Name:", 
    "DBA:", 
    "Name as Shown on License:",
    "Agency License Type:",
    "Agency License Number:",
    "Agency License State:",
    "Agency License Expiration:",
    "Federal Tax ID #:" ]

    pdf = pdfquery.PDFQuery(file)
    pdf.load(0)
    reslt_dict = {}
    for x in genInfoLabel:
        label = pdf.pq('LTTextLineHorizontal:contains("'+ x + '")')
        left_corner = float(label.attr('x0')) + 100
        bottom_corner = float(label.attr('y0'))
        #Coordinates: bottom-left-x, bottom-left-y, top-right-x, top-right-y
        data = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -5, left_corner + 430, bottom_corner +30)).text()
        reslt_dict[x] = data
    
    # for x, y in reslt_dict.items():
    #     print(x[:-1] + ": " + y)
    return reslt_dict

def getAddress(file):
    pdf = pdfquery.PDFQuery(file)
    pdf.load(1)
    label = pdf.pq('LTTextLineHorizontal:contains("Street:")')
    left_corner = float(label.attr('x0')) + 20
    bottom_corner = float(label.attr('y0'))
    street = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -10, left_corner + 430, bottom_corner +30)).text()
    #print("Street: " + street)
    return street

def getProducerCity(file):
    pdf = pdfquery.PDFQuery(file)
    pdf.load(1)
    label = pdf.pq('LTTextLineHorizontal:contains("City:")')
    left_corner = float(label.attr('x0')) + 20
    bottom_corner = float(label.attr('y0'))
    city = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -10, left_corner + 150, bottom_corner +30)).text()
    #print("City: " + city)
    return city

def getProducerState(file):
    pdf = pdfquery.PDFQuery(file)
    pdf.load(1)
    label = pdf.pq('LTTextLineHorizontal:contains("State:")')
    left_corner = float(label.attr('x0')) + 20
    bottom_corner = float(label.attr('y0'))
    state = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -10, left_corner + 100, bottom_corner +30)).text()
    #print("State: " + state)
    return state

def getProducerZip(file):
    pdf = pdfquery.PDFQuery(file)
    pdf.load(1)
    label = pdf.pq('LTTextLineHorizontal:contains("Zip Code:")')
    left_corner = float(label.attr('x0')) + 20
    bottom_corner = float(label.attr('y0'))
    zip = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -10, left_corner + 430, bottom_corner +30)).text()
    #print("Zip Code: " + zip)
    return zip

def getProducerPhone(file):
    pdf = pdfquery.PDFQuery(file)
    pdf.load(1)
    label = pdf.pq('LTTextLineHorizontal:contains("Phone:")')
    left_corner = float(label.attr('x0')) + 20
    bottom_corner = float(label.attr('y0'))
    phone = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -10, left_corner + 100, bottom_corner +30)).text()
    #print("Phone: " + phone)
    return phone

def getProducerFax(file):
    pdf = pdfquery.PDFQuery(file)
    pdf.load(1)
    label = pdf.pq('LTTextLineHorizontal:contains("Fax:")')
    left_corner = float(label.attr('x0')) + 20
    bottom_corner = float(label.attr('y0'))
    fax = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -10, left_corner + 430, bottom_corner +30)).text()
    #print("Fax: " + fax)
    return fax

def writeToFile(args):
    if len(args) == 2:
        file = args[1]
        general_info = getGenInfo(file)
        # Create a Pandas Excel writer using XlsxWriter engine.
        out_path = r"\\files\\"
        file_name = "Producer Import " + general_info["Legal Agency Name:"] + ".xlsx"
        file_new = out_path + file_name
        print("Output will save to: " + str(file_new))
        
        # Create new Excel file and add a worksheet.
        workbook = xlsxwriter.Workbook(file_new)
        worksheet = workbook.add_worksheet()

        # row, column, field
        worksheet.write(0, 0, "ProducerName")
        worksheet.write(1, 0, general_info["Legal Agency Name:"])

        worksheet.write(0, 1, "Name")
        worksheet.write(1, 1, general_info["DBA:"])

        worksheet.write(0, 2, "Address1")
        worksheet.write(1, 2, getAddress(file))

        worksheet.write(0, 3, "City")
        worksheet.write(1, 3, getProducerCity(file))

        worksheet.write(0, 4, "County")
        worksheet.write(1, 4, "")

        worksheet.write(0, 5, "State")
        worksheet.write(1, 5, getProducerState(file))

        worksheet.write(0, 6, "ZipCode")
        worksheet.write(1, 6, getProducerZip(file))

        worksheet.write(0, 7, "Phone")
        worksheet.write(1, 7, getProducerPhone(file))

        worksheet.write(0, 8, "Fax")
        worksheet.write(1, 8, getProducerFax(file))

        worksheet.write(0, 9, "FEIN")
        worksheet.write(1, 9, general_info["Federal Tax ID #:"])

        worksheet.write(0, 10, "Website")
        worksheet.write(1, 10, general_info["Website Address:"])

        worksheet.write(0, 11, "Email")
        worksheet.write(1, 11, general_info["Agency E-mail Address:"])
        
        workbook.close()
    else:
        print("Invalid number of argumants.")

print('Number of arguments:', len(sys.argv), 'arguments.')
print('Argument List:', str(sys.argv))
writeToFile(sys.argv)
