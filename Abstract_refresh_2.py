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



##for y in range (1, 8)



text = Document.Range().Text
