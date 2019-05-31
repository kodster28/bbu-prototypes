import win32com.client as win
import re
import os

word = win.gencache.EnsureDispatch("Word.Application")
word.Visible = False

compiled_abstracts = {}
initial_path = r'C:\Users\kody.jackson\Desktop\Abstracts'
key = 1

for document in os.listdir(initial_path):



    Document = word.Documents.Open(os.path.join(initial_path, document))
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
            document_information[y][1] = re.sub('\\x07', '', document_information[y][1])

            if document_information[y][0] == "COURSE OVERVIEW":
                document_information[y][1] = re.sub('\\r', ' ', document_information[y][1])

            if document_information[y][0] == "DELIVERY METHOD" or document_information[y][0] == "LEARNING OBJECTIVES" or document_information[y][0] == "PREREQUISITES":
                document_information[y][1] = re.split("\\r", document_information[y][1])

                while '' in document_information[y][1]:
                    document_information[y][1].remove('')

                while 'OR' in document_information[y][1]:
                    document_information[y][1].remove('OR')


    compiled_abstracts[key] = document_information
    key += 1

    Document.Close()

word.Quit()


