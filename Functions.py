def split(FullPath):
    from regex import match, findall
    PathOnly = ''
    for find in findall(r'[^\/]+?\/', FullPath):
        PathOnly += find
    FileOnly = match(r'(?r)[^\/]+', FullPath).group()
    return (FullPath.replace('\\', '/'), PathOnly.replace('\\', '/'), FileOnly)


def Upsave(FullPath, PathOnly, FileOnly):
    match FullPath[-3:]:
        case 'doc':
            FullPath = Doc2Docx(FullPath, PathOnly, FileOnly)
            FullPath, PathOnly, FileOnly = split(FullPath)
        case 'ppt':
            FullPath = Ppt2Pptx(FullPath, PathOnly, FileOnly)
            FullPath, PathOnly, FileOnly = split(FullPath)
        case 'xls':
            FullPath = Xls2Xlsx(FullPath, PathOnly, FileOnly)
            FullPath, PathOnly, FileOnly = split(FullPath)
    return (FullPath)


def Doc2Docx(FullPath, PathOnly, FileOnly):
    from win32com.client import DispatchEx

    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath.replace('/', '\\'))
    Doc.SaveAs(PathOnly + FileOnly[:-3] + 'docx', FileFormat=12)
    WordApp.Quit()
    FullPath = PathOnly + FileOnly[:-3] + 'docx'
    return (FullPath)


def Xls2Xlsx(FullPath, PathOnly, FileOnly):
    from win32com.client import DispatchEx

    XlApp = DispatchEx('Excel.Application')
    Xl = XlApp.Workbooks.Open(FullPath.replace('/', '\\'))
    Xl.SaveAs(PathOnly + FileOnly[:-4], FileFormat=51)
    XlApp.Quit()
    FullPath = PathOnly + FileOnly[:-3] + 'xlsx'
    return (FullPath)


def Ppt2Pptx(FullPath, PathOnly, FileOnly):
    from win32com.client import DispatchEx

    PptApp = DispatchEx('PowerPoint.Application')
    Ppt = PptApp.Presentations.Open(FullPath.replace('/', '\\'), 0, 0, 0)
    Ppt.SaveAs(PathOnly + FileOnly[:-4], FileFormat=24)
    PptApp.Quit()
    FullPath = PathOnly + FileOnly[:-3] + 'pptx'
    return (FullPath)


def ExtractImages(FullPath, PathOnly, FileOnly):
    import zipfile as zip
    from regex import match
    from os import makedirs

    file = zip.ZipFile(FullPath)
    makedirs(name=PathOnly + 'Temp', exist_ok=True)
    for media in file.namelist():
        if match(r'(ppt|word|xl|story)/media/.*?\.(jpeg|jpg|png)', media):
            file.extract(media, PathOnly + 'Temp')
    if FileOnly.endswith('pptx') or FileOnly.endswith('.story'):
        for rel in file.namelist():
            if match(r'(ppt|story)/slides/_rels', rel):
                file.extract(rel, PathOnly + 'Temp')


def CleanTempDir(Tempdir):
    from os import walk, remove
    from PIL import Image

    for TempRoot, TempPath, TempFile in walk(Tempdir):
        for file in TempFile:
            if file.endswith('jpeg'):
                im = Image.open(TempRoot.replace('\\', '/') + '/' + file)
                im.save(TempRoot.replace('\\', '/') + '/' + file[:-4] + 'png')
                remove(TempRoot.replace('\\', '/') + '/' + file)
            elif file.endswith('jpg'):
                im = Image.open(TempRoot.replace('\\', '/') + '/' + file)
                im.save(TempRoot.replace('\\', '/') + '/' + file[:-3] + 'png')
                remove(TempRoot.replace('\\', '/') + '/' + file)


def FillCS(Tempdir, PathOnly, FileOnly):
    import docx
    from os import walk, makedirs
    from shutil import rmtree

    makedirs(PathOnly + '\\Contact Sheets', exist_ok=True)
    CS = docx.Document()

    for PicRoot, PicPath, PicFile in walk(Tempdir):
        for pic in PicFile:
            if pic.endswith('png'):
                Table = CS.add_table(rows=5, cols=2, style='Table Grid')
                Table.cell(0, 0).merge(Table.cell(0, 1)).text = pic
                Table.cell(1, 0).merge(Table.cell(1, 1))
                Table.cell(2, 0).merge(Table.cell(2, 1))
                Table.cell(3, 0).text = 'Source'
                Table.cell(3, 1).text = 'Target'
                Table.cell(1, 0).paragraphs[0].add_run().add_picture(
                    PicRoot.replace('\\', '/') + '/' + pic,
                    width=(CS.sections[0].page_width - (CS.sections[0]
                           .right_margin + CS.sections[0].left_margin)))
                if FileOnly.endswith('pptx') or FileOnly.endswith('story'):
                    Locations = LocateImage(Tempdir, pic)
                    for location in Locations:
                        Table.cell(2, 0).add_paragraph(
                            location, style='List Bullet')
                    CS.add_section()
                else:
                    CS.add_section()
    CS.save(PathOnly + 'Contact Sheets/CS_' + FileOnly + '.docx')
    rmtree(Tempdir)


def LocateImage(TempDir, ImageName):
    from os import walk
    from regex import search

    Locations = []
    for Relroot, Relpaths, Relfiles in walk(TempDir):
        for rel in Relfiles:
            if rel.endswith('rels'):
                relstr = str(open(Relroot + '\\' + rel).read())
                if search(ImageName[:-4], relstr):
                    Locations.append(rel[:5] + ' ' + rel[5:rel.find('.')])
    return (Locations)


def BilTables(PathOnly, FileOnly):
    from win32com.client import DispatchEx

    WordApp = DispatchEx('Word.Application')
    WordApp.Visible = False
    Doc = WordApp.Documents.Open(PathOnly.replace('/', '\\') + FileOnly,
                                 Visible=False)
    Doc = WordApp.ActiveDocument
    if Doc.Tables.Count > 0:
        for index, table in enumerate(Doc.Tables):
            yield f'Proccessing table {index + 1} of {Doc.Tables.Count}',\
                   index, Doc.Tables.Count
            for cell in table.Range.Cells:
                for i in range(1, cell.Range.Paragraphs.Count + 1):
                    cell.Range.Paragraphs.Add()
                    cell.Range.Paragraphs(cell.Range.Paragraphs.Count).\
                        Range.FormattedText = cell.Range.Paragraphs(i).\
                        Range.FormattedText
                    cell.Range.Paragraphs(i).Range.Font.Hidden = True

    Doc.SaveAs2(PathOnly + 'Bil_' + FileOnly)
    WordApp.Quit()


def BilText(PathOnly, FileOnly):
    from win32com.client import DispatchEx

    WordApp = DispatchEx('Word.Application')
    WordApp.Visible = False
    Doc = WordApp.Documents.Open(PathOnly.replace('/', '\\') + 'Bil_' +
                                 FileOnly, Visible=False)
    Doc = WordApp.ActiveDocument

    if Doc.Tables.Count > 0:
        for table in Doc.Tables:
            table.Rows.WrapAroundText = True

    i = Doc.Paragraphs.Count
    for index, par in enumerate(Doc.Paragraphs):
        yield f'Proccessing paragraph {index} of {i}', index, i
        if not par.Range.Information(12):
            par.Range.Find.Execute(FindText="^t", ReplaceWith=" ", Replace=2)
            par.Range.Paragraphs.Add(par.Range)
            par.Previous(1).Range.FormattedText = par.Range.FormattedText
            par.Previous(1).Range.Font.Hidden = True
            Doc.Range(par.Previous(1).Range.start, par.Range.End).\
                ConvertToTable(Separator=0, NumRows=1, NumColumns=2)

    Doc.Content.Find.Execute(FindText="^p^p", ReplaceWith="^p", Replace=2)

    for table in Doc.Tables:
        table.Select()
        Doc.Application.Selection.SplitTable()

    for table in Doc.Tables:
        table.Rows.WrapAroundText = False

    Doc.Save()
    WordApp.Quit()


def Doc2PDF(FullPath, PathOnly, FileOnly, ARev, DRev, Com, Overwrite):
    from win32com.client import DispatchEx
    from os import makedirs

    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath.replace('/', '\\'), Visible=False)

    makedirs(PathOnly + 'PDFs', exist_ok=True)
    if Doc.Revisions.Count > 0 and ARev:
        Doc.AcceptAllRevisions()
    if Doc.Revisions.Count > 0 and DRev:
        Doc.RejectAllRevisions()
    if Doc.Comments.Count > 0 and Com:
        Doc.DeleteAllComments()
    if Overwrite:
        Doc.Save()
    Doc.SaveAs(PathOnly.replace('/', '\\') + 'PDFs\\' + FileOnly + r'.pdf',
               FileFormat=17)
    WordApp.Quit()


def AcceptRevisions(FullPath, PathOnly, FileOnly, ARev, DRev, Com,
                    Overwrite):
    from win32com.client import DispatchEx

    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath.replace('/', '\\'))

    if Doc.Revisions.Count > 0 and ARev:
        Doc.AcceptAllRevisions()
    if Doc.Revisions.Count > 0 and DRev:
        Doc.RejectAllRevisions()
    if Doc.Comments.Count > 0 and Com:
        Doc.DeleteAllComments()
    if Overwrite:
        Doc.Save()
    else:
        match FileOnly[-1]:
            case 'x':
                Doc.SaveAs(PathOnly.replace('/', '\\') + 'NoRev_' +
                           FileOnly[:-5], FileFormat=12)
            case 'm':
                Doc.SaveAs(PathOnly.replace('/', '\\') + 'NoRev_' +
                           FileOnly[:-5], FileFormat=12)
            case 'c':
                Doc.SaveAs(PathOnly.replace('/', '\\') + 'NoRev_' +
                           FileOnly[:-4], FileFormat=0)
    WordApp.Quit()


def PrepStoryExport(FullPath):
    from win32com.client import DispatchEx
    from os.path import abspath

    BasPath = abspath(r'StoryWord.bas')
    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath.replace('/', '\\'))
    Doc.VBProject.VBComponents.Import(BasPath)
    Doc.Application.Run('Story')
    WordApp.Quit()


def Unhide(FullPath, PathOnly, FileOnly, ARev, DRev, Com, Row, Col, Sheet,
           Shp, Sld, Overwrite):
    from win32com.client import DispatchEx
    from os.path import abspath

    match FileOnly[-4:]:
        case 'docx' | '.doc' | 'docm':
            BasPath = abspath(r'UnhideWord.bas')
            WordApp = DispatchEx('Word.Application')
            Doc = WordApp.Documents.Open(FullPath.replace('/', '\\'))
            Doc.VBProject.VBComponents.Import(BasPath)
            if Doc.Revisions.Count > 0 and ARev:
                Doc.AcceptAllRevisions()
            if Doc.Revisions.Count > 0 and DRev:
                Doc.RejectAllRevisions()
            if Com and Doc.Comments.Count > 0:
                Doc.DeleteAllComments()
            Doc.Application.Run('Unhide')
            if Overwrite:
                Doc.Save()
            else:
                match FileOnly[-1]:
                    case 'x':
                        Doc.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                   FileOnly[:-5], FileFormat=12)
                    case 'm':
                        Doc.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                   FileOnly[:-5], FileFormat=13)
                    case 'c':
                        Doc.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                   FileOnly[:-4], FileFormat=0)
            Doc.VBProject.VBComponents.Remove(
                Doc.VBProject.VBComponents.Item("PrepToolKit1"))
            WordApp.Quit()
        case 'xlsx' | '.xls' | 'xlsm':
            BasPath = abspath(r'UnhideExcel.bas')
            XlApp = DispatchEx('Excel.Application')
            Xl = XlApp.Workbooks.Open(FullPath.replace('/', '\\'))
            Xl.VBProject.VBComponents.Import(BasPath)
            if Sheet:
                Xl.Application.Run('UnhideSheet')
            if Row:
                Xl.Application.Run('UnhideRow')
            if Col:
                Xl.Application.Run('UnhideCol')
            if Overwrite:
                Xl.Save()
            else:
                match FileOnly[-1]:
                    case 'x':
                        Xl.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                  FileOnly[:-5],
                                  FileFormat=51)
                    case 'm':
                        Xl.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                  FileOnly[:-5],
                                  FileFormat=52)
                    case 's':
                        Xl.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                  FileOnly[:-4],
                                  FileFormat=43)
            Xl.VBProject.VBComponents.Remove(
                Xl.VBProject.VBComponents.Item("PrepToolKit1"))
            XlApp.Quit()
        case 'pptx' | '.ppt' | 'pptm':
            BasPath = abspath(r'UnhidePPT.bas')
            PptApp = DispatchEx('PowerPoint.Application')
            Ppt = PptApp.Presentations.Open(FullPath.replace('/', '\\'),
                                            0, 0, 0)
            Ppt.VBProject.VBComponents.Import(BasPath)
            if Sld:
                Ppt.Application.Run('UnhideSlide')
            if Shp:
                Ppt.Application.Run('UnhideShape')
            if Overwrite:
                Ppt.Save()
            else:
                match FileOnly[-1]:
                    case 'x':
                        Xl.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                  FileOnly[:-5],
                                  FileFormat=24)
                    case 'm':
                        Xl.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                  FileOnly[:-5],
                                  FileFormat=25)
                    case 't':
                        Xl.SaveAs(PathOnly.replace('/', '\\') + 'UNH_' +
                                  FileOnly[:-4],
                                  FileFormat=1)
            Xl.VBProject.VBComponents.Remove(
                Xl.VBProject.VBComponents.Item("PrepToolKit1"))
            XlApp.Quit()
