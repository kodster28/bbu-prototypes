import win32com.client, sys
from LMS_file_creator_Excel import create_LMS_folder

filename = "C:\\Users\\kody.jackson\\Documents\\Stats_test\\Stat_test_2.pptx"
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True



Presentation = Application.Presentations.Open(filename)
sourceCount = 0
slideList = []
sourceList = ["Blackbaud Charitable Giving Report", "Charitable Giving Report", "charitable giving report"]

for x in range(1, Presentation.Slides.Count + 1):


    for shape in range(1, Presentation.Slides(x).Shapes.Count + 1):


        if Presentation.Slides(x).Shapes(shape).HasTextFrame or "TextBox" in Presentation.Slides(x).Shapes(shape).Name:

            ##add in try and except blocks

            if any(word in Presentation.Slides(x).Shapes(shape).TextFrame.TextRange.Text for word in sourceList):
                sourceCount += 1
                slideList.append(Presentation.Slides(x).SlideNumber)

print(filename)

for item in slideList:
    print(str(item))


##Presentation.Close()
##
##Application.Quit()
