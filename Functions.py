def Upsave(FullPath):
	from regex import match, findall

	PathOnly = ''
	for find in findall(r'[^\/]+?\/',FullPath):
		PathOnly += find
	FileOnly = match(r'(?r)[^\/]+',FullPath).group()

	match FullPath[-3:]:
		case 'doc':
			FullPath = Doc2Docx(FullPath, PathOnly, FileOnly)
		case 'ppt':
			FullPath = Ppt2Pptx(FullPath, PathOnly, FileOnly)
		case 'xls':
			FullPath = Xls2Xlsx(FullPath, PathOnly, FileOnly)
	return(FullPath, PathOnly, FileOnly)

def Doc2Docx(FullPath, PathOnly, FileOnly):
	from win32com.client import Dispatch

	WordApp = Dispatch('Word.Application')
	Doc = WordApp.Documents.Open(FullPath)
	Doc.SaveAs(PathOnly + FileOnly[:-4], FileFormat=12)
	WordApp.Quit()
	FullPath = PathOnly + FileOnly[:-3] + 'docx'
	return(FullPath)

def Xls2Xlsx(FullPath, PathOnly, FileOnly):
	from win32com.client import Dispatch

	XlApp = Dispatch('Excel.Application')
	Xl = XlApp.Workbooks.Open(FullPath)
	Xl.SaveAs(PathOnly + FileOnly[:-4], FileFormat=51)
	XlApp.Quit()
	return(PathOnly + FileOnly[:-3] + 'xlsx')

def Ppt2Pptx(FullPath, PathOnly, FileOnly):
	from win32com.client import Dispatch

	PptApp = Dispatch('PowerPoint.Application')
	Ppt = PptApp.Presentations.Open(FullPath,0,0,0)
	Ppt.SaveAs(PathOnly + FileOnly[:-4], FileFormat=24)
	PptApp.Quit()
	return(PathOnly + FileOnly[:-3]+ 'pptx')

def ExtractImages(FullPath, PathOnly, FileOnly):
	import zipfile as zip
	from regex import match
	from os import makedirs,rename

	file = zip.ZipFile(FullPath)
	makedirs(name=PathOnly + 'Temp',exist_ok=True)
	for media in file.namelist():
		if match(r'(ppt|word|xl)/media', media):
			file.extract(media,PathOnly + 'Temp')
	CleanTempDir(PathOnly + 'Temp')		
	# for rel in file.namelist():
	# 	if match(r'ppt/slides/_rels', rel):
	# 		file.extract(rel,PathOnly + 'Temp')

def CleanTempDir(Tempdir):
	from os import walk, remove
	from PIL import Image

	for TempRoot, TempPath, TempFile in walk(Tempdir):
		for file in TempFile:
			if file[-4:] == 'jpeg':
				with Image.open(TempRoot.replace('\\','/') + '/' + file) as im:
					im.save(TempRoot.replace('\\','/') + '/' + file[:-4] + 'png')
					remove(TempRoot.replace('\\','/') + '/' + file)
			elif file[-3:] == 'jpg':
				with Image.open(TempRoot.replace('\\','/') + '/' + file) as im:
					im.save(TempRoot.replace('\\','/') + '/' + file[:-3] + 'png')
					remove(TempRoot.replace('\\','/') + '/' + file)
			elif file[-3:] == 'png':
				continue
			else:
				remove(TempRoot.replace('\\','/') + '/' + file)

def FillCS(Tempdir, PathOnly, FileOnly):
	import docx
	from os import walk, makedirs
	from shutil import rmtree

	makedirs(PathOnly + 'Contact Sheets', exist_ok=True)
	CS = docx.Document()

	for PicRoot, PicPath, PicFile in walk(Tempdir):
		for pic in PicFile:
			Table = CS.add_table(rows=5, cols=2, style='Table Grid')
			Table.cell(0,0).merge(Table.cell(0,1)).text = pic
			Table.cell(1,0).merge(Table.cell(1,1))
			Table.cell(2,0).merge(Table.cell(2,1))
			Table.cell(3,0).text = 'Source'
			Table.cell(3,1).text = 'Target'
			Table.cell(1,0).add_paragraph().add_run().add_picture(PicRoot.replace('\\', '/') + '/' + pic, width=(CS.sections[0].page_width - (CS.sections[0].right_margin + CS.sections[0].left_margin)))
			CS.add_section()
	CS.save(PathOnly + 'Contact Sheets/CS_' + FileOnly + '.docx')
	rmtree(Tempdir)