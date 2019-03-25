import os
import win32com.client, sys
folder_path = 'C:\\Users\\kody.jackson\\Desktop\\python_test'

Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True

for item in os.listdir(folder_path):
    if ".pptx" in item:
        filename = os.path.join(folder_path, item)
        
        
        Presentation = Application.Presentations.Open(filename)


        nameList = ["database view", "Database view", "Database View"]
        notesCount = 0
        slideCount = 0

        for x in range(1, Presentation.Slides.Count + 1):
            
            for shape in range(1, Presentation.Slides(x).Shapes.Count + 1):
                
                if Presentation.Slides(x).Shapes(shape).HasTextFrame or "TextBox" in Presentation.Slides(x).Shapes(shape).Name:
                      
                      if any(word in Presentation.Slides(x).Shapes(shape).TextFrame.TextRange.Text for word in nameList):
                          slideCount += 1
                          Presentation.Slides(x).Comments.Add(Left= 12, Top= 12, Author= "Kody", AuthorInitials= "KJ", Text= "This slide has the phrase on its slide face")

                if "Group" in Presentation.Slides(x).Shapes(shape).Name:
                    
                    for items in range(1, Presentation.Slides(x).Shapes(shape).GroupItems.Count + 1):
                        
                        if Presentation.Slides(x).Shapes(shape).GroupItems(items).HasTextFrame or "TextBox" in Presentation.Slides(x).Shapes(shape).GroupItems(items).Name:
                            if any(word in Presentation.Slides(x).Shapes(shape).GroupItems(items).TextFrame.TextRange.Text for word in nameList):
                                slideCount += 1
                                Presentation.Slides(x).Comments.Add(Left= 12, Top= 12, Author= "Kody", AuthorInitials= "KJ", Text= "This slide has the phrase on its slide face")
                        
            for shape in range(1, Presentation.Slides(x).NotesPage.Shapes.Count + 1):
                if Presentation.Slides(x).NotesPage.Shapes(shape).HasTextFrame:
                    if any(word in Presentation.Slides(x).NotesPage.Shapes(shape).TextFrame.TextRange.Text for word in nameList):
                        notesCount += 1
                        Presentation.Slides(x).Comments.Add(Left= 12, Top= 12, Author= "Kody", AuthorInitials= "KJ", Text= "The slide text has been touched")
                        while any(word in Presentation.Slides(x).NotesPage.Shapes(shape).TextFrame.TextRange.Text for word in nameList):
                            Presentation.Slides(x).NotesPage.Shapes(shape).TextFrame.TextRange.Replace(FindWhat= "database view", ReplaceWhat= "advanced administration view", MatchCase = True)
                            Presentation.Slides(x).NotesPage.Shapes(shape).TextFrame.TextRange.Replace(FindWhat= "Database view", ReplaceWhat= "Advanced administration view", MatchCase = True)
                            Presentation.Slides(x).NotesPage.Shapes(shape).TextFrame.TextRange.Replace(FindWhat= "Database View", ReplaceWhat= "Advanced Administration View", MatchCase = True)



        Presentation.Slides(1).Comments.Add(Left= 12, Top= 12, Author= "Kody", AuthorInitials= "KJ", Text= "The program has addressed notes issues on " + str(notesCount) + ' number of slides.')

        Presentation.Slides(1).Comments.Add(Left= 12, Top= 12, Author= "Kody", AuthorInitials= "KJ", Text= "The program has called out " + str(slideCount) + ' issues in the slide face.')

        if notesCount > 0 or slideCount > 0:
            Presentation.SaveAs(os.path.join(folder_path, 'Items_Found_Needs_Verification', item))
        else:
            Presentation.SaveAs(os.path.join(folder_path, 'None_Found', item))
            
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Quit()


