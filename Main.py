import PySimpleGUI as gui
import Functions as fn
from configparser import ConfigParser
from os.path import isfile

config = ConfigParser()
config.read('config.ini')

TopText = [
	[gui.Text(text='Welcome to Max GUI!\nThis utility lets you run the different functions of Max (also knows as the Watched Folders) anywhere you want.\nSimply choose a function from the sidebar on the left, click on Browse, select all the files you want to process and click on run.',justification='center')],
	[gui.HorizontalSeparator()]
]

Sidebar = [
	[gui.Button(button_text='Contact Sheet',size=15)],
	[gui.Button(button_text='Bilingual Table',size=15)],
	[gui.Button(button_text='ChExcel',size=15)],
	[gui.Button(button_text='Other',size=15)]
]

MainCanvas = [
	[gui.InputText(default_text='',size=75,key='---PathInput---'),gui.FilesBrowse(target='---PathInput---')],
	[gui.Text(text='',key='---Description---',size=(48,5))],
	[gui.Submit(button_text='Run',size=15,visible=False,key='---Run---')],
	[gui.ProgressBar(max_value=100,orientation='horizontal',size=(48,20),bar_color=('green','white'),key='---PBar---',visible=False)],
	[gui.Text(text='',key='---PBarFile---')],
	[gui.Text(text='',key='---PBarFileStep---')]
]

layout = [[TopText, gui.Column(Sidebar, vertical_scroll_only=True, scrollable=True,expand_y=True), gui.VSeparator(), gui.Column(MainCanvas)]]

MainWindow = gui.Window('Max GUI', layout)

while True:
	event, values = MainWindow.read()
	if event == 'Exit' or event == gui.WIN_CLOSED:
		break
	if event == '---Run---':
		PathList = values['---PathInput---'].split(';')
		MainWindow['---PBar---'].update(visible=True)
		for index, PathInput in enumerate(PathList):
			FullPath, PathOnly, FileOnly = fn.Upsave(PathInput)
			MainWindow['---PBarFile---'].update(FileOnly)
			MainWindow['---PBarFileStep---'].update('Extracting Images')
			MainWindow['---PBar---'].update((index+1/3)/len(PathList)*100)
			MainWindow.refresh()
			fn.ExtractImages(FullPath, PathOnly, FileOnly)
			MainWindow['---PBarFileStep---'].update('Filling in Contact Sheet')
			MainWindow['---PBar---'].update((index+2/3)/len(PathList)*100)
			MainWindow.refresh()
			fn.FillCS(PathOnly + 'Temp', PathOnly, FileOnly)
			MainWindow['---PBar---'].update((index+3/3)/len(PathList)*100)
			MainWindow.refresh()
		MainWindow['---PBar---'].update(100)
		MainWindow['---PBarFile---'].update('')
		MainWindow['---PBarFileStep---'].update('Done!')
	else:
		if event == 'Contact Sheet':
			MainWindow['---Description---'].update(config['Descriptions']['Contact_Sheet'])
			MainWindow['---Run---'].update(visible=True)
		if event == 'Bilingual Table':
			MainWindow['---Description---'].update(config['Descriptions']['Bilingual_Table'])
			MainWindow['---Run---'].update(visible=True)

MainWindow.close()