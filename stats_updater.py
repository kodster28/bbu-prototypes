import win32com.client, sys

filename = "C:\\Users\\kody.jackson\\Documents\\Stats_test\\Stat_test_2.pptx"
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True


Presentation = Application.Presentations.Open(filename)
Slides = Presentation.Slides

stats_dict = {"22%" : "45%", "29%": "74%"}


old_source = "Blackbaud Charitable Giving Report, 2018"
new_source = "Blackbaud Charitable Giving Report, 2019"

for x in range(1, Slides.Count + 1):

    try:

        if old_source in Slides(x).Shapes("Source_Text").TextFrame.TextRange.Text:

            print("Got the slide")
                
            for shape in range(1, Slides(x).Shapes.Count + 1):

                print("This is shape " + Slides(x).Shapes(shape).Name)
                    
                if Slides(x).Shapes(shape).HasTextFrame or "TextBox" in Slides(x).Shapes(shape).Name:


                    print("I have text")

                    for key in stats_dict:

                        if key in Slides(x).Shapes(shape).TextFrame.TextRange.Text:


                            Slides(x).Shapes("Source_Text").TextFrame.TextRange.Replace(FindWhat = old_source, ReplaceWhat = new_source, MatchCase = True)
                            Slides(x).Shapes(shape).TextFrame.TextRange.Replace(FindWhat= key, ReplaceWhat= stats_dict[key], MatchCase = True)
                            Slides(x).Comments.Add(Left= 12, Top= 12, Author= "Kody", AuthorInitials= "KJ", Text= "I replaced a stat here")
                            print("is this working?")
                        


                    


    except:

        print("This slide doesn't have that value")


                
##Presentation.Save()
##Presentation.Close()
##
##Application.Quit()


##can't search for the textbox individually, but can loop through the slide to search for the source name
