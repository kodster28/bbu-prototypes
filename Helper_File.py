def check_shape_for_text(shape):
    #returns true if shape in MS Office has text. Requires win32com.client module.
    if shape.HasTextFrame or 'TextBox' in shape.Name:
        return True
    else:
        return False

def is_shape_a_group(shape):
    #returns true if shape in MS Office is part of a group. Requires win32com.client module.
    try:
        if shape.GroupItems.Count > 0:
            return True
    except:
        return False

def find_replace(shape, find_phrase, replace_phrase):
    #finds and replaces text within a shape in MS Office. Requires win32com.client module.
    return shape.TextFrame.TextRange.Replace(FindWhat=find_phrase, ReplaceWhat=replace_phrase, MatchCase = True)

def add_comment(object, author, text):
    #adds a comment to a PPT or Word Document. Make sure you select slide object for PPT, document object for Word. Does not work for Excel. Requires win32com.client module.
    return object.Comments.Add(Left = 12, Top = 12, Author = author, AuthorInitials = 'Py', Text = text)

def open_office_application(win32_import, application_type, visible):
    #opens an office application. Requires proper naming of objects (Word, Excel, PowerPoint). Returns application object. Visible is boolean.
    Application = win32_import.gencache.EnsureDispatch(application_type + '.Application')
    Application.Visible = visible
    return Application

