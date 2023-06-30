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
	from win32com.client import DispatchEx

	WordApp = DispatchEx('Word.Application')
	Doc = WordApp.Documents.Open(FullPath)
	Doc.SaveAs(PathOnly + FileOnly[:-4], FileFormat=12)
	WordApp.Quit()
	FullPath = PathOnly + FileOnly[:-3] + 'docx'
	return(FullPath)

def Xls2Xlsx(FullPath, PathOnly, FileOnly):
	from win32com.client import DispatchEx

	XlApp = DispatchEx('Excel.Application')
	Xl = XlApp.Workbooks.Open(FullPath)
	Xl.SaveAs(PathOnly + FileOnly[:-4], FileFormat=51)
	XlApp.Quit()
	return(PathOnly + FileOnly[:-3] + 'xlsx')

def Ppt2Pptx(FullPath, PathOnly, FileOnly):
	from win32com.client import DispatchEx

	PptApp = DispatchEx('PowerPoint.Application')
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
		if match(r'(ppt|word|xl|story)/media/.*?\.(jpeg|jpg|png)', media):
			file.extract(media,PathOnly + 'Temp')
	CleanTempDir(PathOnly + 'Temp')
	if FileOnly.endswith('pptx') or FileOnly.endswith('.story'):
		for rel in file.namelist():
			if match(r'(ppt|story)/slides/_rels', rel):
				file.extract(rel,PathOnly + 'Temp')

def CleanTempDir(Tempdir):
	from os import walk, remove
	from PIL import Image

	for TempRoot, TempPath, TempFile in walk(Tempdir):
		for file in TempFile:
			if file.endswith('jpeg'):
				with Image.open(TempRoot.replace('\\','/') + '/' + file) as im:
					im.save(TempRoot.replace('\\','/') + '/' + file[:-4] + 'png')
					remove(TempRoot.replace('\\','/') + '/' + file)
			elif file.endswith('jpg'):
				with Image.open(TempRoot.replace('\\','/') + '/' + file) as im:
					im.save(TempRoot.replace('\\','/') + '/' + file[:-3] + 'png')
					remove(TempRoot.replace('\\','/') + '/' + file)

def FillCS(Tempdir, PathOnly, FileOnly):
	import docx
	from os import walk, makedirs
	from shutil import rmtree
	from regex import match
	makedirs(PathOnly + 'Contact Sheets', exist_ok=True)
	CS = docx.Document()

	for PicRoot, PicPath, PicFile in walk(Tempdir):
		for pic in PicFile:
			if pic.endswith('png'):
				Table = CS.add_table(rows=5, cols=2, style='Table Grid')
				Table.cell(0,0).merge(Table.cell(0,1)).text = pic
				Table.cell(1,0).merge(Table.cell(1,1))
				Table.cell(2,0).merge(Table.cell(2,1))
				Table.cell(3,0).text = 'Source'
				Table.cell(3,1).text = 'Target'
				Table.cell(1,0).paragraphs[0].add_run().add_picture(PicRoot.replace('\\', '/') + '/' + pic, width=(CS.sections[0].page_width - (CS.sections[0].right_margin + CS.sections[0].left_margin)))
				if FileOnly.endswith('pptx') or FileOnly.endswith('story'):
					Locations = LocateImage(Tempdir, pic)
					for location in Locations:
						Table.cell(2,0).add_paragraph(location, style='List Bullet')
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
	return(Locations)