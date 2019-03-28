import win32com.client as win
import re

word = win.gencache.EnsureDispatch("Word.Application")
word.Visible = False
Document = word.Documents.Open("c://users//kody.jackson//desktop//LO_Optimizing_Your_Email_Performance_Abstract_011519.docx")
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


with open('C:\\Users\\kody.jackson\\Desktop\\' + document_information[0][1] + '_' + document_information[1][1] + '.txt', 'w') as document:
        for z in range(1, 10):
                document.write('%s\n' % document_information[z])


Document.Close()
word.Quit()
