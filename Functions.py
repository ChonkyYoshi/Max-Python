import docx
import openpyxl as xl
from openpyxl.utils.cell import get_column_letter
import pptx
import regex as re
import os
import win32com.client as com
import zipfile as zip
from PIL import Image
import shutil
import helper as hp


def split(FullPath):
    PathOnly = ''
    for find in re.findall(r'[^\/]+?\/', FullPath):
        PathOnly += find
    FileOnly = re.match(r'(?r)[^\/]+', FullPath).group()
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

    WordApp = com.DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath.replace('/', '\\'))
    Doc.SaveAs(PathOnly + FileOnly[:-3] + 'docx', FileFormat=12)
    WordApp.Quit()
    FullPath = PathOnly + FileOnly[:-3] + 'docx'
    return (FullPath)


def Xls2Xlsx(FullPath, PathOnly, FileOnly):

    XlApp = com.DispatchEx('Excel.Application')
    Xl = XlApp.Workbooks.Open(FullPath.replace('/', '\\'))
    Xl.SaveAs(PathOnly + FileOnly[:-4], FileFormat=51)
    XlApp.Quit()
    FullPath = PathOnly + FileOnly[:-3] + 'xlsx'
    return (FullPath)


def Ppt2Pptx(FullPath, PathOnly, FileOnly):

    PptApp = com.DispatchEx('PowerPoint.Application')
    Ppt = PptApp.Presentations.Open(FullPath.replace('/', '\\'), 0, 0, 0)
    Ppt.SaveAs(PathOnly + FileOnly[:-4], FileFormat=24)
    PptApp.Quit()
    FullPath = PathOnly + FileOnly[:-3] + 'pptx'
    return (FullPath)


def ExtractImages(FullPath, PathOnly, FileOnly):

    file = zip.ZipFile(FullPath)
    os.makedirs(name=PathOnly + 'Temp', exist_ok=True)
    for media in file.namelist():
        if re.match(r'(ppt|word|xl|story)/media/.*?\.(jpeg|jpg|png)', media):
            file.extract(media, PathOnly + 'Temp')
    if FileOnly.endswith('pptx') or FileOnly.endswith('.story'):
        for rel in file.namelist():
            if re.match(r'(ppt|story)/slides/_rels', rel):
                file.extract(rel, PathOnly + 'Temp')


def CleanTempDir(Tempdir):

    for TempRoot, TempPath, TempFile in os.walk(Tempdir):
        for file in TempFile:
            if file.endswith('jpeg'):
                im = Image.open(TempRoot.replace('\\', '/') + '/' + file)
                im.save(TempRoot.replace('\\', '/') + '/' + file[:-4] + 'png')
                os.remove(TempRoot.replace('\\', '/') + '/' + file)
            elif file.endswith('jpg'):
                im = Image.open(TempRoot.replace('\\', '/') + '/' + file)
                im.save(TempRoot.replace('\\', '/') + '/' + file[:-3] + 'png')
                os.remove(TempRoot.replace('\\', '/') + '/' + file)


def FillCS(Tempdir, PathOnly, FileOnly):

    os.makedirs(PathOnly + '\\Contact Sheets', exist_ok=True)
    CS = docx.Document()

    for PicRoot, PicPath, PicFile in os.walk(Tempdir):
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
    shutil.rmtree(Tempdir)


def LocateImage(TempDir, ImageName):

    Locations = []
    for Relroot, Relpaths, Relfiles in os.walk(TempDir):
        for rel in Relfiles:
            if rel.endswith('rels'):
                relstr = str(open(Relroot + '\\' + rel).read())
                if re.search(ImageName[:-4], relstr):
                    Locations.append(rel[:5] + ' ' + rel[5:rel.find('.')])
    return (Locations)


def BilTable(PathOnly, FileOnly):

    li = list()
    doc = docx.Document(PathOnly + FileOnly)
    for index, table in enumerate(doc.tables):
        yield f'processing table {index + 1} of {len(doc.tables)}',\
            index/len(doc.tables)
        for c in range(len(table.columns)):
            for r in range(len(table.rows)):
                if table.cell(r, c)._tc not in li:
                    li.append(table.cell(r, c)._tc)
                    for par in table.cell(r, c).paragraphs:
                        prevpar = par.insert_paragraph_before()
                        hp.CopyParFormatting(prevpar, par)
                        for run in par.runs:
                            prevrun = prevpar.add_run()
                            prevrun.text = run.text
                            hp.CopyRunFormatting(prevrun, run)
                            prevrun.font.hidden = True
        li.clear()
    i = len(doc.paragraphs)
    for index, par in enumerate(doc.paragraphs):
        yield f'processing paragraph {index + 1} of {i}', index/i
        table = doc.add_table(rows=1, cols=2)
        par._p.addnext(table._tbl)
        SPar = table.cell(0, 0).paragraphs[0]
        TPar = table.cell(0, 1).paragraphs[0]
        hp.CopyParFormatting(SPar, par)
        hp.CopyParFormatting(TPar, par)
        for run in par.runs:
            SRun = SPar.add_run()
            TRun = TPar.add_run()
            SRun.text = run.text
            TRun.text = run.text
            hp.CopyRunFormatting(SRun, run)
            hp.CopyRunFormatting(TRun, run)
        par._element.getparent().os.remove(par._element)
    doc.save(PathOnly + 'Bil_' + FileOnly)


def Doc2PDF(FullPath, PathOnly, FileOnly, ARev, DRev, Com, Overwrite):

    WordApp = com.DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(FullPath.replace('/', '\\'), Visible=False)

    os.makedirs(PathOnly + 'PDFs', exist_ok=True)
    if Doc.Revisions.Count > 0 and ARev:
        Doc.AcceptAllRevisions()
    if Doc.Revisions.Count > 0 and DRev:
        Doc.RejectAllRevisions()
    if Doc.Comments.Count > 0 and Com:
        Doc.DeleteAllComments()
    if Overwrite:
        Doc.Save()
    Doc.SaveAs2(PathOnly.replace('/', '\\') + 'PDFs\\' + FileOnly + r'.pdf',
                FileFormat=17)
    WordApp.Quit()


def AcceptRevisions(FullPath, PathOnly, FileOnly, ARev, DRev, Com,
                    Overwrite):

    WordApp = com.DispatchEx('Word.Application')
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
                Doc.SaveAs2(PathOnly.replace('/', '\\') + 'NoRev_' +
                            FileOnly, FileFormat=12)
            case 'm':
                Doc.SaveAs2(PathOnly.replace('/', '\\') + 'NoRev_' +
                            FileOnly, FileFormat=13)
            case 'c':
                Doc.SaveAs2(PathOnly.replace('/', '\\') + 'NoRev_' +
                            FileOnly[:-4], FileFormat=0)
    WordApp.Quit()


def PrepStoryExport(FullPath, PathOnly, FileOnly, Regex):

    Doc = docx.Document(FullPath)
    Regex = '(?i)' + Regex
    for index, par in enumerate(Doc.paragraphs):
        yield f'paragraph {index + 1} of {len(Doc.paragraphs)}',\
            index/len(Doc.paragraphs)
        for run in par.runs:
            run.font.hidden = True
    for table in Doc.tables:
        for index, col in enumerate(table.columns):
            yield f'column {index + 1} of {len(table.columns)}',\
                index/len(table.columns)
            if not index == 3:
                for cell in col.cells:
                    for par in cell.paragraphs:
                        for run in par.runs:
                            run.font.hidden = True
            else:
                for cell in col.cells:
                    for par in cell.paragraphs:
                        for run in par.runs:
                            if re.match(Regex, run.text):
                                start = re.match(Regex, run.text).start()
                                end = re.match(Regex, run.text).end()
                                hidden_run = hp.isolate_run(par, start, end)
                                hidden_run.font.hidden = True
                for par in table.cell(0, 3).paragraphs:
                    for run in par.runs:
                        run.font.hidden = True
    Doc.save(PathOnly + 'Prep_' + FileOnly)


def Unhide(FullPath, PathOnly, FileOnly, Row, Col, Sheet,
           Shp, Sld, Overwrite):

    match FileOnly[-4:]:
        case 'docx' | 'docm':
            Doc = docx.Document(FullPath)
            for index, par in enumerate(Doc.paragraphs):
                yield f'Paragraph {index + 1} of {len(Doc.Paragraphs)}',\
                    index/len(Doc.Paragraphs)
                for run in par.runs:
                    run.font.hidden = False
            for index, table in enumerate(Doc.tables):
                yield f'Table {index + 1} of {len(Doc.tables)}',\
                    index/len(Doc.tables)
                for row in table.rows:
                    for cell in row.cells:
                        for par in cell.paragraphs:
                            for run in par.runs:
                                run.font.hidden = False
            for section in Doc.sections:
                for par in section.header.paragraphs:
                    for run in par.runs:
                        run.font.hidden = False
                for par in section.footer.paragraphs:
                    for run in par.runs:
                        run.font.hidden = False
            if Overwrite:
                Doc.save(FullPath)
            else:
                Doc.save(PathOnly + 'UNH_' + FileOnly)
        case 'xlsx' | 'xlsm':
            wb = xl.load_workbook(filename=FullPath)
            for index, ws in enumerate(wb.worksheets):
                yield f'Sheet {ws.title}', index/len(wb.worksheets)
                if not Sheet:
                    ws.sheet_state = 'visible'
                if not Row:
                    for row in range(1, ws.max_row + 1):
                        ws.row_dimensions[row].hidden = False
                if not Col:
                    for col in range(1, ws.max_column + 1):
                        col = get_column_letter(col)
                        ws.column_dimensions[col].hidden = False
            if not Overwrite:
                wb.save()
            else:
                wb.save(PathOnly + 'UNH_' + FileOnly)
        case 'pptx' | 'pptm':
            Pres = pptx.Presentation(FullPath)
            for index, slide in enumerate(Pres.slides):
                yield f'Slide {index} of {len(Pres.Slides)}',\
                    index/len(Pres.slides)
                if not Sld:
                    slide._element.set('show', '1')
                if not Shp:
                    for shape in slide.shapes:
                        shape._element.nvSpPr.cNvPr.set('hidden', '0')
            if Overwrite:
                Pres.save(FullPath)
            else:
                Pres.save(PathOnly + 'UNH_' + FileOnly)
