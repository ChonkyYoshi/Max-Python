def split(FullPath):
    from regex import match, findall
    PathOnly = ''
    for find in findall(r'[^\/]+?\/', FullPath):
        PathOnly += find
    FileOnly = match(r'(?r)[^\/]+', FullPath).group()
    return (FullPath.replace('/', '\\'), PathOnly.replace('/', '\\'), FileOnly)


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
    Doc = WordApp.Documents.Open(FullPath)
    Doc.SaveAs(PathOnly + FileOnly[:-3] + 'docx', FileFormat=12)
    WordApp.Quit()
    FullPath = PathOnly + FileOnly[:-3] + 'docx'
    return (FullPath)


def Doc2PDF(FullPath, PathOnly, FileOnly, AccTC, RejTC, DelCom, Overwrite):
    from win32com.client import DispatchEx
    from os import makedirs

    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath)

    makedirs(PathOnly + 'PDFs', exist_ok=True)
    if Doc.Revisions.Count > 0 and AccTC:
        Doc.AcceptAllRevisions()
    elif Doc.Revisions.Count > 0 and RejTC:
        Doc.RejectAllRevisions()
    if Doc.Comments.Count > 0 and DelCom:
        Doc.DeleteAllComments()
    if Overwrite:
        Doc.Save()
    else:
        Doc.SaveAs(PathOnly + 'PDFs/' + FileOnly + r'.pdf', FileFormat=17)
    WordApp.Quit()


def Xls2Xlsx(FullPath, PathOnly, FileOnly):
    from win32com.client import DispatchEx

    XlApp = DispatchEx('Excel.Application')
    Xl = XlApp.Workbooks.Open(FullPath)
    Xl.SaveAs(PathOnly + FileOnly[:-4], FileFormat=51)
    XlApp.Quit()
    FullPath = PathOnly + FileOnly[:-3] + 'xlsx'
    return (FullPath)


def Ppt2Pptx(FullPath, PathOnly, FileOnly):
    from win32com.client import DispatchEx

    PptApp = DispatchEx('PowerPoint.Application')
    Ppt = PptApp.Presentations.Open(FullPath, 0, 0, 0)
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


def BilTables(FullPath, PathOnly, FileOnly, BasPath):
    from win32com.client import DispatchEx

    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath)
    Doc.VBProject.VBComponents.Import(BasPath)
    Doc.Application.Run('Bil_Tables')
    WordApp.Quit()
    return (PathOnly + '\\Temp_' + FileOnly)


def BilText(FullPath, BasPath):
    from win32com.client import DispatchEx

    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath)
    Doc.VBProject.VBComponents.Import(BasPath)
    Doc.Application.Run('Bil_Text')
    WordApp.Quit()


def AcceptRevisions(FullPath, PathOnly, FileOnly, AccTC, RejTC, DelCom,
                    Overwrite):
    from win32com.client import DispatchEx

    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath)

    if Doc.Revisions.Count > 0 and AccTC:
        Doc.AcceptAllRevisions()
    elif Doc.Revisions.Count > 0 and RejTC:
        Doc.RejectAllRevisions()
    if Doc.Comments.Count > 0 and DelCom:
        Doc.DeleteAllComments()
    if Overwrite:
        Doc.Save()
    else:
        match FileOnly[-1]:
            case 'x':
                Doc.SaveAs(PathOnly + 'NoRev_' + FileOnly[:-5], FileFormat=12)
            case 'm':
                Doc.SaveAs(PathOnly + 'NoRev_' + FileOnly[:-5], FileFormat=12)
            case 'c':
                Doc.SaveAs(PathOnly + 'NoRev_' + FileOnly[:-4], FileFormat=0)
    WordApp.Quit()


def PrepStoryExport(FullPath, BasPath):
    from win32com.client import DispatchEx

    print(FullPath)
    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath)
    Doc.VBProject.VBComponents.Import(BasPath)
    Doc.Application.Run('Story')
    WordApp.Quit()
