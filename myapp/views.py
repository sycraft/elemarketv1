import os
import openpyxl
from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from django.template import RequestContext

def index(request):
    if "GET" == request.method:
        wb = openpyxl.load_workbook('.\\resources\\temp.xlsx')
        sheets = wb.sheetnames
        active_sheet = wb.active
        excel_data = list()
        # iterating over the rows and
        # getting value from each cell in row
        worksheet = wb["Sheet1"]
        for row in worksheet.iter_rows():
            row_data = list()
            for cell in row:
                row_data.append(str(cell.value))
                # print(cell.value)
            excel_data.append(row_data)
        return render(request, 'myapp/index.html', {"excel_data":excel_data})
    else:
        excel_file = request.FILES["excel_file"]
        destination = open(os.path.join(".\\resources", 'temp.xlsx'), 'wb+')
        for chunk in excel_file.chunks():
            destination.write(chunk)
        destination.close()
        # you may put validations here to check extension or file size

        wb = openpyxl.load_workbook(excel_file)

        # getting all sheets
        sheets = wb.sheetnames
        # print(sheets)

        # getting a particular sheet
        worksheet = wb["Sheet1"]
        # print(worksheet)

        # getting active sheet
        active_sheet = wb.active
        # print(active_sheet)

        # reading a cell
        # print(worksheet["A1"].value)

        excel_data = list()
        # iterating over the rows and
        # getting value from each cell in row
        for row in worksheet.iter_rows():
            row_data = list()
            for cell in row:
                row_data.append(str(cell.value))
                # print(cell.value)
            excel_data.append(row_data)

        return render(request, 'myapp/index.html', {"excel_data":excel_data})

def download(request):
    file = open('.\\resources\\temp.xlsx', 'rb')
    response = HttpResponse(file)
    response['Content-Type'] = 'application/octet-stream'
    response['Content-Disposition'] = 'attachment;filename="temp.xlsx'
    return response