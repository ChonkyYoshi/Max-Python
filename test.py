import win32com.client


def search_replace_all(word_file, find_str, replace_str):
    wdFindContinue = 1
    wdReplaceAll = 2

    # Dispatch() attempts to do a GetObject() before creating a new one.
    # DispatchEx() just creates a new one.
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = 0
    app.DisplayAlerts = 0
    app.Documents.Open(word_file)

    # expression.Execute(FindText, MatchCase, MatchWholeWord,
    #   MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward,
    #   Wrap, Format, ReplaceWith, Replace)
    app.Selection.Find.Execute(find_str, False, False, False, False, False,
                               True, wdFindContinue, False, replace_str,
                               wdReplaceAll)
    app.ActiveDocument.Close(SaveChanges=True)
    app.Quit()


f = 'C:\\Users\\eagosta\\Downloads\\New folder (69)\\' +\
    'WF_wco_templates_test yourself mod 7.docx'
search_replace_all(f, 'Story', 'Test')
