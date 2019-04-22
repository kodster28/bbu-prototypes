import win32com.client as win
import os

##set up the folder path, variables
old_folder_path = "C:\\Users\\kody.jackson\\Desktop\\datasheets\\Old"
new_folder_path = "C:\\Users\\kody.jackson\\Desktop\\datasheets\\Changed_Links"
course_name = "OBP: Basics of Events"
new_hyperlink = "https://www.google.com"

##launch Word instance
word = win.gencache.EnsureDispatch("Word.Application")
word.Visible = False

for item in os.listdir(old_folder_path):

    filename = os.path.join(old_folder_path, item)
    word.Visible = False
    document = word.Documents.Open(filename)
    document.TrackRevisions = True

    selectionRange = document.Content

    if selectionRange.Find.Execute(FindText= course_name):
        selectionRange.Hyperlinks(1).Address = new_hyperlink

        elements = item.split("_")
        elements[-1] = elements[-1].split('.')
        elements[-1][0] = elements[-1][0] + '_updated'
        elements[-1] = '.'.join(elements[-1])
        new_file_name = '_'.join(elements)

        document.SaveAs(os.path.join(new_folder_path, new_file_name))
        document.Close()
    else:
        document.Close()

word.Quit()
