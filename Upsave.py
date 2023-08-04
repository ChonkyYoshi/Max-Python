from pathlib import Path
from win32com.client import DispatchEx


def Upsave(File: Path):
    match File.suffix:
        case '.doc':
            File = Doc2Docx(File)
        case '.ppt':
            File = Ppt2Pptx(File)
        case '.xls':
            File = Xls2Xlsx(File)
    return File


def Doc2Docx(File: Path):

    WordApp = DispatchEx('Word.Application')
    Doc = WordApp.Documents.Open(File.__str__())
    Doc.SaveAs(f'{File.__str__()}.docx', FileFormat=12)
    WordApp.Quit()
    File = Path(f'{File.__str__()}.docx')
    return File


def Xls2Xlsx(File: Path):

    XlApp = DispatchEx('Excel.Application')
    Xl = XlApp.Workbooks.Open(File.__str__())
    Xl.SaveAs(File.__str__(), FileFormat=51)
    XlApp.Quit()
    File = Path(f'{File.__str__()}.xlsx')
    return File


def Ppt2Pptx(File: Path):

    PptApp = DispatchEx('PowerPoint.Application')
    Ppt = PptApp.Presentations.Open(File.__str__(), 0, 0, 0)
    Ppt.SaveAs(File.__str__(), FileFormat=24)
    PptApp.Quit()
    File = Path(f'{File.__str__()}.pptx')
    return File
