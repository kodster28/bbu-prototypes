##For saving files in the new LMS format

import os, shutil, win32com.client

#define the function

def create_LMS_folder (excel_row):

    ##declare variables, copy, rename, and add the target folder to the CSOD
    course_name = sheet.Cells(excel_row, 1).Value
    course_file_name = sheet.Cells(excel_row, 2).Value
    product_abbreviation = sheet.Cells(excel_row, 3).Value
    email_address = sheet.Cells(excel_row, 4).Value
    new_folder_path = os.path.join(folder_path, product_abbreviation, course_file_name)

    shutil.copytree(os.path.join(folder_path, "Generic_Class_Resources"), new_folder_path)

    ##rename the individual files
    os.rename(os.path.join(new_folder_path, "Generic_Abstract.pdf"), os.path.join(new_folder_path, course_file_name + "_Abstract.pdf"))
    os.rename(os.path.join(new_folder_path, "Generic_Handout.pdf"), os.path.join(new_folder_path, course_file_name + "_Handout.pdf"))
    os.rename(os.path.join(new_folder_path, "Generic_Resources.docx"), os.path.join(new_folder_path, course_file_name + "_Resources.docx"))
    resource_file = os.path.join(new_folder_path, course_file_name + "_Resources")


    ##open the word file and perform tasks
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    document = word.Documents.Open(resource_file + ".docx")

    ##find and replace course name
    find = word.Selection.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Text = "Generic Course Name"
    find.Replacement.Text = course_name
    find.Forward = True
    find.Wrap = win32com.client.constants.wdFindContinue
    find.Execute(Replace=win32com.client.constants.wdReplaceAll)

    ##add hyperlink. Make sure that begins with http
    hyperlink = "https://www.blackbaud.com/files/training/jobaids/CSOD/" + product_abbreviation + '/' + course_file_name + '/' + course_file_name + '_Handout.pdf'
    document.Hyperlinks.Add(document.InlineShapes(3).Range, hyperlink)

    ##save and save as PDF
    document.Save()
    document.ExportAsFixedFormat(os.path.join(resource_file + '.pdf'), 17)
    document.Close()
    word.Quit()

    return;


##call the function for all values in an excel file
folder_path = "C:\\Users\\kody.jackson\\Desktop\\CSOD_test"
excel_file = os.path.join(folder_path, "LMS_file_input.xlsx")

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
workbook = xl.Workbooks.Open(excel_file)
sheet = workbook.ActiveSheet


row = 2
while sheet.Cells(row, 1).Value != None:
    create_LMS_folder(row)
    row += 1

workbook.Save()
workbook.Close()
xl.Quit()
