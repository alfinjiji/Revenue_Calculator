from flask import Flask, render_template, redirect, url_for, request
from werkzeug.utils import secure_filename
import os, openpyxl, secrets, xlsxwriter
import os.path as op
from pathlib import Path

app = Flask(__name__)

def save_file(exl_file, fl):
    random_hex = secrets.token_hex(4)
    _, f_ext = os.path.splitext(exl_file.filename)
    sheet = fl + random_hex + f_ext
    exl_file.save('InputFile/' +sheet )
    return sheet

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/revenue', methods=['POST'])
def revenue():
    if request.method == 'POST':
        # Read excel sheet
        f1 = request.files['sheet1']
        fl1 = "1sh"
        sh1 = save_file(f1, fl1)
        
        f2 = request.files['sheet2']
        fl2 = "2sh"
        sh2 = save_file(f2, fl2)
        # Path of input folder 
        path1 =  op.join(op.dirname(__file__), 'InputFile')
        # Path of output folder 
        path2 =  op.join(op.dirname(__file__), 'OutputFile')
        # Read Excel Sheet 1 (Project Details)
        xlsx_file1 = Path(path1, sh1)
        wb_obj1 = openpyxl.load_workbook(xlsx_file1)
        sheet1 = wb_obj1.active
        # Read Excel Sheet 2 (Employee Timesheet)
        xlsx_file2 = Path(path1, sh2)
        wb_obj2 = openpyxl.load_workbook(xlsx_file2)
        sheet2 = wb_obj2.active

        # Generate excel sheet
        # sh3 = output file name
        sh3 = secrets.token_hex(3) + "_report.xlsx"
        workbook = xlsxwriter.Workbook('OutputFile/'+sh3) 
        worksheet = workbook.add_worksheet() 
        # write data to excel
        worksheet.write('A1', 'Project Name') 
        worksheet.write('B1', 'Project Estimation (in billed currency)')
        worksheet.write('C1', 'Other Expenses')
        worksheet.write('D1', 'Employees worked')
        worksheet.write('E1', 'Rate/Day (INR)')
        worksheet.write('F1', 'No. of Hours worked') 
        worksheet.write('G1', 'Expected Revenue (for each employee)') 
        worksheet.write('H1', 'Actual Revenue')
        worksheet.write('I1', 'Profit')
        worksheet.write('J1', 'Loss')
    
        x = 2
        ar = 0
        # demo start
        r = []
        for i in range(2, sheet1.max_row):
            cell = sheet1.cell(row=i, column=1).value
            if cell != None:
                r.append(cell)
        # demo end

        for i in range(2, sheet1.max_row):
            my_cell1 = sheet1.cell(row=i, column=1)
            if my_cell1.value != None:
                prj = my_cell1.value
                prj = prj.replace(" ", "")
                prj = prj.lower()
            for j in range(2, sheet2.max_row):
                prj2 = sheet2.cell(row=j, column=2).value
                if prj2 != None:
                    prj2 = prj2.replace(" ", "")
                    prj2 = prj2.lower()
                
                # if employee name of sheet1 equal to employee name of sheet2 and project name of sheet1 equal to project name of sheet2 
                if sheet1.cell(row=i, column=3).value == sheet2.cell(row=j, column=1).value and  prj2 == prj:
                    # project name
                    worksheet.write("A"+str(x), prj)
                    # project estimation
                    worksheet.write("B"+str(x), sheet1.cell(row=i, column=2).value)
                    # other expense
                    worksheet.write("C"+str(x), sheet1.cell(row=i, column=5).value)
                    # employee
                    worksheet.write("D"+str(x), sheet1.cell(row=i, column=3).value)
                    # rate per day
                    worksheet.write("E"+str(x), sheet1.cell(row=i, column=4).value)
                    # no. of hour worked
                    worksheet.write("F"+str(x), sheet2.cell(row=j, column=3).value)
                    # calculating expected revenue
                    rate = sheet1.cell(row=i, column=4).value
                    worked = sheet2.cell(row=j, column=3).value
                    expected_revenue = float(rate * (worked/8))
                    worksheet.write("G"+str(x), expected_revenue)
                    x = x+1
        workbook.close()
        # Achual Revenue calculation
        xlsx_file3 = Path(path2, sh3)
        wb_obj3 = openpyxl.load_workbook(xlsx_file3)
        sheet3 = wb_obj3.active
        k=0
        ind=2
        for i in range(2, sheet3.max_row+1):
            cell1 = sheet3.cell(row=i, column=1)
            cell7 = sheet3.cell(row=i, column=7)
            if i == 2:
                k=k+cell7.value
            elif cell1.value == sheet3.cell(row=i-1, column=1).value and i != 2:
                k=k+cell7.value
            else:
                print(ind)
                H=sheet3.cell(row=ind, column=8)
                H.value = k
                ind=i
                print(k)
                k=cell7.value
            if i==sheet3.max_row:
                print(ind)
                H=sheet3.cell(row=ind, column=8)
                H.value = k
                print(k)
        ach_rev = os.path.join(path2, sh3)
        wb_obj3.save(ach_rev)
    return str(ach_rev)
    #return redirect(url_for('index'))

if __name__ =='__main__':
    app.run(debug=True) 