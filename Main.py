import PySimpleGUI as gui
import Functions as fn
from configparser import ConfigParser
from os.path import isfile, abspath
from os import remove

config = ConfigParser()
config.read('config.ini')
filestypes = dict(config['file_types'])
for key in filestypes:
    filestypes[key] = tuple(str(filestypes[key]).split(','))


def Contact_Sheet(PathInput):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    if FileOnly[-3:] in ['doc', 'ppt', 'xls']:
        MainWindow['---PBarFile---'].update(FileOnly)
        MainWindow['---PBarFileStep---'].update('Upsaving to Office 2007 ' +
                                                'format')
        MainWindow['---PBar---'].update((index+1/5)/len(PathList)*100)
        FullPath = fn.Upsave(FullPath, PathOnly, FileOnly)
    MainWindow['---PBarFile---'].update(FileOnly)
    MainWindow['---PBarFileStep---'].update('Extracting Images')
    MainWindow['---PBar---'].update((index+2/5)/len(PathList)*100)
    MainWindow.refresh()
    fn.ExtractImages(FullPath, PathOnly, FileOnly)
    MainWindow['---PBarFileStep---'].update('cleaning up jpeg and jpg')
    MainWindow['---PBar---'].update((index+3/5)/len(PathList)*100)
    MainWindow.refresh()
    fn.CleanTempDir(PathOnly + 'Temp')
    MainWindow['---PBarFileStep---'].update('Filling in Contact Sheet')
    MainWindow['---PBar---'].update((index+4/5)/len(PathList)*100)
    MainWindow.refresh()
    fn.FillCS(PathOnly + 'Temp', PathOnly, FileOnly)


def Bilingual(PathInput):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    if FileOnly[-3:] in ['doc', 'ppt', 'xls']:
        MainWindow['---PBarFile---'].update(FileOnly)
        MainWindow['---PBarFileStep---'].update('Upsaving to Office 2007\
                                                 format')
        MainWindow['---PBar---'].update((index+1/5)/len(PathList)*100)
        FullPath = fn.Upsave(FullPath, PathOnly, FileOnly)
    MainWindow['---PBarFile---'].update(FileOnly)
    MainWindow['---PBarFileStep---'].update('Upsaving to Office 2007 format')
    MainWindow['---PBar---'].update((index+1/5)/len(PathList)*100)
    FullPath = fn.Upsave(FullPath)
    MainWindow['---PBarFile---'].update(FileOnly)
    MainWindow['---PBarFileStep---'].update('Processing tables')
    MainWindow['---PBar---'].update((index + 2/5)/len(PathList)*100)
    FullPath = fn.BilTables(FullPath.replace('/', '\\\\'), PathOnly, FileOnly,
                            abspath(r'Bas Files\\Bil.bas').
                            replace('\\', '\\\\'))
    MainWindow.refresh()
    MainWindow['---PBarFile---'].update(FileOnly)
    MainWindow['---PBarFileStep---'].update('Processing regular text')
    MainWindow['---PBar---'].update((index + 3/5)/len(PathList)*100)
    fn.BilText(FullPath.replace('/', '\\\\'),
               abspath(r'Bas Files\\Bil.bas').replace('\\', '\\\\'))
    MainWindow['---PBarFile---'].update(FileOnly)
    MainWindow['---PBarFileStep---'].update('Removing Temp file')
    MainWindow['---PBar---'].update((index + 4/5)/len(PathList)*100)
    MainWindow.refresh()
    remove(FullPath)


def Doc2PDF(PathInput, AccTC, RejTC, DelCom, Overwrite):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    MainWindow['---PBarFile---'].update(FileOnly)
    MainWindow['---PBarFileStep---'].update('Saving as PDF')
    MainWindow['---PBar---'].update(index/len(PathList)*100)
    FullPath = fn.Doc2PDF(FullPath, PathOnly, FileOnly, AccTC, RejTC, DelCom,
                          Overwrite)


def AcceptRevisions(PathInput, AccTC, RejTC, DelCom, Overwrite):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    MainWindow['---PBarFile---'].update(FileOnly)
    MainWindow['---PBarFileStep---'].update('Accepting Revisions')
    MainWindow['---PBar---'].update(index/len(PathList)*100)
    fn.AcceptRevisions(FullPath, PathOnly, FileOnly, AccTC, RejTC,
                       DelCom, Overwrite)


def PrepStoryExport(PathInput):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    print(FullPath)
    MainWindow['---PBarFile---'].update(FileOnly)
    MainWindow['---PBarFileStep---'].update('Prepping files')
    MainWindow['---PBar---'].update(index/len(PathList)*100)
    fn.PrepStoryExport(FullPath, abspath(r'Bas Files\\Story.bas').
                       replace('\\', '\\\\'))


TopText = [
    [gui.Text(text=str(config['Descriptions']['TopText']),
              justification='center')],
    [gui.HorizontalSeparator()]
]

Sidebar = [
    [gui.Button(button_text='Contact Sheet', size=15)],
    [gui.Button(button_text='Bilingual Table', size=15)],
    [gui.Button(button_text='Word to PDF', size=15)],
    [gui.Button(button_text='Accept Revisions', size=15)],
    [gui.Button(button_text='Prep Story Export', size=15)]
]

Rightlayout = [
    [gui.InputText(default_text='', key='---PathInput---',
                   visible=False),
     gui.FilesBrowse(target='---PathInput---',
                     visible=False, key='---Browse---')],
    [gui.Text(text='', key='---Description---')],
    [gui.Submit(button_text='Run', size=15, visible=False, key='---Run---')],
    [gui.Checkbox(text='Accept Revisions', auto_size_text=True,
     key='---AcceptTC---', visible=False),
     gui.Checkbox(text='Reject Revisions',
     auto_size_text=True, key='---RejectTC---', visible=False)],
    [gui.Checkbox(text='Delete Comments', auto_size_text=True,
     key='---DeleteCom---', visible=False)],
    [gui.Checkbox(text='Overwrite', auto_size_text=True, key='---Overwrite---',
     visible=False)],
    [gui.ProgressBar(max_value=100, orientation='horizontal', size=(50, 20),
     bar_color=('green', 'white'), key='---PBar---', visible=False)],
    [gui.Text(text='', key='---PBarFile---')],
    [gui.Text(text='', key='---PBarFileStep---')]
]

layout = [[TopText, gui.Column(Sidebar, vertical_scroll_only=True,
                               scrollable=True, expand_y=True),
          gui.VSeparator(),
          gui.Column(Rightlayout)]]

MainWindow = gui.Window('Prep ToolKit', layout, size=(780, 400))
FunctionName = ''

while True:
    event, values = MainWindow.read()
    match event:
        case 'Exit' | gui.WIN_CLOSED:
            break
        case '---Run---':
            if MainWindow['---AcceptTC---'].get() is True and\
               MainWindow['---RejectTC---'].get() is True:
                gui.popup_error('Both Accept and Reject TC are selected!\n\
                Please check only of the options and try again.',
                                title='Both accept and reject selected')
                break
            PathList = values['---PathInput---'].split(';')
            MainWindow['---PBar---'].update(visible=True)

            for index, PathInput in enumerate(PathList):
                if not isfile(PathInput):
                    gui.popup_error(f'the following file:\n{PathInput}\n\
                    is not a valid file!\n Skipping this file',
                                    title='Invalid file error',
                                    auto_close=True, auto_close_duration=5,
                                    keep_on_top=True, modal=True)
                    continue
                match FunctionName:
                    case 'Contact_Sheet':
                        Contact_Sheet(PathInput)
                    case 'Bilingual_Table':
                        Bilingual(PathInput)
                    case 'Doc2PDF':
                        Doc2PDF(PathInput, MainWindow['---AcceptTC---'].get(),
                                MainWindow['---RejectTC---'].get(),
                                MainWindow['---DeleteCom---'].get(),
                                MainWindow['---Overwrite---'].get())
                    case 'Accept Revisions':
                        AcceptRevisions(PathInput,
                                        MainWindow['---AcceptTC---'].get(),
                                        MainWindow['---RejectTC---'].get(),
                                        MainWindow['---DeleteCom---'].get(),
                                        MainWindow['---Overwrite---'].get())
                    case 'Prep_Story':
                        PrepStoryExport(PathInput)
            MainWindow['---PBar---'].update(100)
            MainWindow['---PBarFile---'].update('')
            MainWindow['---PBarFileStep---'].update('Done!')
        case 'Contact Sheet':
            MainWindow['---PathInput---'].update(visible=True)
            MainWindow['---Browse---'].update(visible=True)
            MainWindow['---Browse---'].FileTypes = filestypes['cs'],
            MainWindow['---Description---'].update(config['Descriptions']
                                                   ['Contact_Sheet'])
            MainWindow['---Run---'].update(visible=True)
            MainWindow['---AcceptTC---'].update(visible=False)
            MainWindow['---RejectTC---'].update(visible=False)
            MainWindow['---DeleteCom---'].update(visible=False)
            MainWindow['---Overwrite---'].update(visible=False)
            FunctionName = 'Contact_Sheet'
        case 'Bilingual Table':
            MainWindow['---PathInput---'].update(visible=True)
            MainWindow['---Browse---'].update(visible=True)
            MainWindow['---Browse---'].FileTypes = filestypes['bil'],
            MainWindow['---Description---'].update(config['Descriptions']
                                                   ['Bilingual_Table'])
            MainWindow['---Run---'].update(visible=True)
            MainWindow['---AcceptTC---'].update(visible=False)
            MainWindow['---RejectTC---'].update(visible=False)
            MainWindow['---DeleteCom---'].update(visible=False)
            MainWindow['---Overwrite---'].update(visible=False)
            FunctionName = 'Bilingual_Table'
        case 'Word to PDF':
            MainWindow['---PathInput---'].update(visible=True)
            MainWindow['---Browse---'].update(visible=True)
            MainWindow['---Browse---'].FileTypes = filestypes['pdf'],
            MainWindow['---Description---'].update(config['Descriptions']
                                                   ['Doc2PDF'])
            MainWindow['---Run---'].update(visible=True)
            MainWindow['---AcceptTC---'].update(visible=True)
            MainWindow['---RejectTC---'].update(visible=True)
            MainWindow['---DeleteCom---'].update(visible=True)
            MainWindow['---Overwrite---'].update(visible=True)
            FunctionName = 'Doc2PDF'
        case 'Accept Revisions':
            MainWindow['---PathInput---'].update(visible=True)
            MainWindow['---Browse---'].update(visible=True)
            MainWindow['---Browse---'].FileTypes = filestypes['rev'],
            MainWindow['---Description---'].update(config['Descriptions']
                                                   ['Accept_Revisions'])
            MainWindow['---Run---'].update(visible=True)
            MainWindow['---AcceptTC---'].update(visible=True)
            MainWindow['---RejectTC---'].update(visible=True)
            MainWindow['---DeleteCom---'].update(visible=True)
            MainWindow['---Overwrite---'].update(visible=True)
            FunctionName = 'Accept Revisions'
        case 'Prep Story Export':
            MainWindow['---PathInput---'].update(visible=True)
            MainWindow['---Browse---'].update(visible=True)
            MainWindow['---Browse---'].FileTypes = filestypes['story'],
            MainWindow['---Description---'].update(config['Descriptions']
                                                   ['Prep_Story'])
            MainWindow['---Run---'].update(visible=True)
            MainWindow['---AcceptTC---'].update(visible=False)
            MainWindow['---RejectTC---'].update(visible=False)
            MainWindow['---DeleteCom---'].update(visible=False)
            MainWindow['---Overwrite---'].update(visible=False)
            FunctionName = 'Prep_Story'
MainWindow.close()
