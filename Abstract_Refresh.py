import win32com.client as win
import re

word = win.gencache.EnsureDispatch("Word.Application")
word.Visible = True
Document = word.Documents.Open("c://users//kody.jackson//desktop//RENXT_FUN_Fundraising_Abstract_120518.docx")
header = Document.Sections.Item(1).Headers(1)

document_information = [['PRODUCT NAME', ''],
                        ['COURSE NAME', ''],
                        ['COURSE TYPE', ''],
                        ['MODALITY', ''],
                        ['DURATION', ''],
                        ['DELIVERY METHOD', ''],
                        ['PREREQUISITES', ''],
                        ['COURSE OVERVIEW', ''],
                        ['TARGET AUDIENCE', ''],
                        ['LEARNING OBJECTIVES', ''],
                        ['VIEW ADDITIONAL INFORMATION', '']]




for x in range (1, header.Shapes.Count + 1):
	try:
		if header.Shapes(x).TextFrame.TextRange.Font.Bold:
			document_information[0][1] = header.Shapes(x).TextFrame.TextRange.Text.strip()

		else:
			document_information[1][1] = header.Shapes(x).TextFrame.TextRange.Text.strip()
	except:
		print ()

text = Document.Range().Text


for y in range (2, 10):
        nameRegex = re.compile(document_information[y][0]+ '(.*?)' + document_information[y + 1][0])
        document_information[y][1] = nameRegex.search(text).group(1).strip()



##with open('C:\\Users\\kody.jackson\\Desktop\\' + document_information[0][1] + '_' + document_information[1][1] + '.txt', 'w') as document:
##        for z in range(0, 10):
##                document.write('%s\n' % document_information[z])



Document.Close()

Document = word.Documents.Open('C:\\Users\\kody.jackson\\Desktop\\Old_Abstract.dotx')
##Document = word.Documents.SaveAs('C:\\Users\\kody.jackson\\Desktop\\' + document_information[0][1] + '_' + document_information[1][1] + '.docx')

##find and replace through the body of the document

for z in range (0, len(document_information)):
    if ' ' in document_information[z][0]:
        document_information[z][0] = document_information[z][0].lower().capitalize().replace(' ', '_')
    else:
        document_information[z][0] = '_' + document_information[z][0].lower() + '_'


for c in range (2, len(document_information) - 1):
    if len(document_information[c][1]) < 255:
        find = word.Selection.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = document_information[c][0]
        find.Replacement.Text = document_information[c][1]
        find.Forward = True
        find.Wrap = win.constants.wdFindContinue
        find.Execute(Replace=win.constants.wdReplaceAll)

    else:
        start = 0
        stop = 200
        step = 200
        length_to_go = True
        while length_to_go:

            if len(document_information[c][1]) < start:
                find = word.Selection.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                find.Text = document_information[c][0]
                find.Replacement.Text = ''
                find.Forward = True
                find.Wrap = win.constants.wdFindContinue
                find.Execute(Replace=win.constants.wdReplaceAll)
                length_to_go = False

            else:
                find = word.Selection.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                find.Text = document_information[c][0]
                find.Replacement.Text = document_information[c][1][start:stop] + document_information[c][0]
                find.Forward = True
                find.Wrap = win.constants.wdFindContinue
                find.Execute(Replace=win.constants.wdReplaceAll)
                start += step
                stop += step


##word.Quit()
