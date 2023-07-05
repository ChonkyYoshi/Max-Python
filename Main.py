import PySimpleGUI as gui
import Functions as fn
from configparser import ConfigParser
from os.path import isfile
from os import remove

Break = False
config = ConfigParser()
config.read('config.ini')
file_types = dict(config['file_types'])
for key in file_types:
    file_types[key] = tuple(file_types[key].split(','))


def ClearOptions():
    for element in MainWindow.element_list():
        if element.metadata == 'option':
            element.update(text='')
            element.update(visible=False)


def SetOptions(Function):
    match Function:
        case 'Contact_Sheet':
            ClearOptions()
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(config['Descriptions']
                                             ['Contact_Sheet'])
        case 'Bilingual_Table':
            ClearOptions()
            MainWindow['R1O1'].update(text='Accept Revisions')
            MainWindow['R1O1'].update(visible=True)
            MainWindow['R2O1'].update(text='Reject Revisions')
            MainWindow['R2O1'].update(visible=True)
            MainWindow['R3O1'].update(text='Delete Comments')
            MainWindow['R3O1'].update(visible=True)
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(config['Descriptions']
                                             ['Bilingual_Table'])
        case 'Doc2PDF':
            ClearOptions()
            MainWindow['R1O1'].update(text='Accept Revisions')
            MainWindow['R1O1'].update(visible=True)
            MainWindow['R2O1'].update(text='Reject Revisions')
            MainWindow['R2O1'].update(visible=True)
            MainWindow['R3O1'].update(text='Delete Comments')
            MainWindow['R3O1'].update(visible=True)
            MainWindow['R2O1'].update(text='Overwrite')
            MainWindow['R2O1'].update(visible=True)
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(config['Descriptions']['Doc2PDF'])
        case 'Accept_Revisions':
            ClearOptions()
            MainWindow['R1O1'].update(text='Accept Revisions')
            MainWindow['R1O1'].update(visible=True)
            MainWindow['R2O1'].update(text='Reject Revisions')
            MainWindow['R2O1'].update(visible=True)
            MainWindow['R3O1'].update(text='Delete Comments')
            MainWindow['R3O1'].update(visible=True)
            MainWindow['R1O2'].update(text='Overwrite')
            MainWindow['R1O2'].update(visible=True)
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(config['Descriptions']
                                             ['Accept_Revisions'])
        case 'Prep_Story':
            ClearOptions()
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(config['Descriptions']
                                             ['Prep_Story'])
        case 'Unhide':
            ClearOptions()
            MainWindow['R1O1'].update(text='Word: Accept Revisions')
            MainWindow['R1O1'].update(visible=True)
            MainWindow['R2O1'].update(text='Word: Reject Revisions')
            MainWindow['R2O1'].update(visible=True)
            MainWindow['R3O1'].update(text='Word: Delete Comments')
            MainWindow['R3O1'].update(visible=True)
            MainWindow['R1O2'].update(text='Excel: Unhide all rows')
            MainWindow['R1O2'].update(visible=True)
            MainWindow['R2O2'].update(text='Excel: Unhide all columns')
            MainWindow['R2O2'].update(visible=True)
            MainWindow['R3O2'].update(text='Excel: Unhide all sheets')
            MainWindow['R3O2'].update(visible=True)
            MainWindow['R1O3'].update(text='Powerpoint: Unhide all shapes')
            MainWindow['R1O3'].update(visible=True)
            MainWindow['R2O3'].update(text='Powerpoint: Unhide all slides')
            MainWindow['R2O3'].update(visible=True)
            MainWindow['R3O3'].update(text='Global: Overwrite')
            MainWindow['R3O3'].update(visible=True)
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(config['Descriptions']
                                             ['Unhide'])
    return Function


def Collapsible(layout, key, title='', arrows=(gui.SYMBOL_DOWN, gui.SYMBOL_UP),
                collapsed=False):
    return gui.Column([[gui.T((arrows[1] if collapsed else arrows[0]),
                      enable_events=True, k=key+'-BUTTON-'), gui.T(title,
                      enable_events=True, key=key+'-TITLE-')],
                      [gui.pin(gui.Column(layout, key=key,
                       visible=not collapsed, metadata=arrows))]], pad=(0, 0))


def Contact_Sheet(PathInput):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    if FileOnly[-3:] in ['doc', 'ppt', 'xls']:
        MainWindow['PBarFile'].update(FileOnly)
        MainWindow['PBarFileStep'].update('Upsaving to Office 2007 ' +
                                          'format...')
        MainWindow['PBar'].update((index+1/5)/len(PathList)*100)
        FullPath = fn.Upsave(FullPath, PathOnly, FileOnly)
    MainWindow['PBarFile'].update(FileOnly)
    MainWindow['PBarFileStep'].update('Extracting Images...')
    MainWindow['PBar'].update((index+2/5)/len(PathList)*100)
    MainWindow.refresh()
    fn.ExtractImages(FullPath, PathOnly, FileOnly)
    MainWindow['PBarFileStep'].update('cleaning up jpeg and jpg...')
    MainWindow['PBar'].update((index+3/5)/len(PathList)*100)
    MainWindow.refresh()
    fn.CleanTempDir(PathOnly.replace('\\', '/') + 'Temp')
    MainWindow['PBarFileStep'].update('Filling in Contact Sheet...')
    MainWindow['PBar'].update((index+4/5)/len(PathList)*100)
    MainWindow.refresh()
    fn.FillCS(PathOnly + 'Temp', PathOnly, FileOnly)


def Bilingual(PathInput):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    if FileOnly[-3:] in ['doc', 'ppt', 'xls']:
        MainWindow['PBarFile'].update(FileOnly)
        MainWindow['PBarFileStep'].update('Upsaving to Office 2007\
                                                 format...')
        MainWindow['PBar'].update((index+1/5)/len(PathList)*100)
        FullPath = fn.Upsave(FullPath, PathOnly, FileOnly)
    MainWindow['PBarFile'].update(FileOnly)
    MainWindow['PBarFileStep'].update('Upsaving to Office 2007 format...')
    MainWindow['PBar'].update((index+1/5)/len(PathList)*100)
    FullPath = fn.Upsave(FullPath)
    MainWindow['PBarFile'].update(FileOnly)
    MainWindow['PBarFileStep'].update('Processing tables...')
    MainWindow['PBar'].update((index + 2/5)/len(PathList)*100)
    FullPath = fn.BilTables(FullPath.replace('/', '\\\\'), PathOnly, FileOnly,)
    MainWindow.refresh()
    MainWindow['PBarFile'].update(FileOnly)
    MainWindow['PBarFileStep'].update('Processing regular text...')
    MainWindow['PBar'].update((index + 3/5)/len(PathList)*100)
    fn.BilText(FullPath.replace('/', '\\\\'))
    MainWindow['PBarFile'].update(FileOnly)
    MainWindow['PBarFileStep'].update('Removing Temp file...')
    MainWindow['PBar'].update((index + 4/5)/len(PathList)*100)
    MainWindow.refresh()
    remove(FullPath)


def Doc2PDF(PathInput, AccTC, RejTC, DelCom, Overwrite):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    MainWindow['PBarFile'].update(FileOnly)
    MainWindow['PBarFileStep'].update('Saving as PDF...')
    MainWindow['PBar'].update(index/len(PathList)*100)
    FullPath = fn.Doc2PDF(FullPath, PathOnly, FileOnly, AccTC, RejTC, DelCom,
                          Overwrite)


def AcceptRevisions(PathInput, AccTC, RejTC, DelCom, Overwrite):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    MainWindow['PBarFile'].update(FileOnly)
    MainWindow['PBarFileStep'].update('Accepting revisions...')
    MainWindow['PBar'].update(index/len(PathList)*100)
    fn.AcceptRevisions(FullPath, PathOnly, FileOnly, AccTC, RejTC,
                       DelCom, Overwrite)


def PrepStoryExport(PathInput):
    FullPath, PathOnly, FileOnly = fn.split(PathInput)
    MainWindow['PBarFile'].update(FileOnly)
    MainWindow['PBarFileStep'].update('Prepping exports...')
    MainWindow['PBar'].update(index/len(PathList)*100)
    fn.PrepStoryExport(FullPath)


def Unhide(PathInput):
    if MainWindow['R1O1'].get() is True and MainWindow['R2O1'].get() is True:
        gui.popup_error('Both Accept and Reject revisions are ticked!\n\
        Please choose only one and try again', title='impossible options',
                        modal=True)
        global Break
        Break = True
    else:
        FullPath, PathOnly, FileOnly = fn.split(PathInput)
        MainWindow['PBarFile'].update(FileOnly)
        MainWindow['PBarFileStep'].update('Unhiding...')
        fn.Unhide(FullPath, PathOnly, FileOnly, Rev=MainWindow['R1O1'].get(),
                  Com=MainWindow['R3O1'].get(), Row=MainWindow['R1O2'].get(),
                  Col=MainWindow['R2O2'].get(), Sheet=MainWindow['R3O2'].get(),
                  Shp=MainWindow['R3O1'].get(), Sld=MainWindow['R3O2'].get(),
                  Overwrite=MainWindow['R3O3'].get())


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
    [gui.Button(button_text='Prep Story Export', size=15)],
    [gui.Button(button_text='Unhide', size=15)]
]

Options = [
    [gui.Checkbox(text='', auto_size_text=True,
     key='R1O1', visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='',
     auto_size_text=True, key='R1O2',
     visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='',
     auto_size_text=True, key='R1O3',
     visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='',
     auto_size_text=True, key='R1O4',
     visible=False, size=(22, 1), pad=(0, 0), metadata='option')],
    [gui.Checkbox(text='', auto_size_text=True,
     key='R2O1', visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='', auto_size_text=True,
     key='R2O2', visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='', auto_size_text=True,
     key='R2O3', visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='', auto_size_text=True,
     key='R2O4', visible=False, size=(22, 1), pad=(0, 0), metadata='option')],
    [gui.Checkbox(text='', auto_size_text=True,
     key='R3O1', visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='', auto_size_text=True,
     key='R3O2', visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='', auto_size_text=True,
     key='R3O3', visible=False, size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='', auto_size_text=True,
     key='R3O4', visible=False, size=(22, 1), pad=(0, 0), metadata='option')]
]

Rightlayout = [
    [gui.InputText(default_text='', key='PathInput',
                   visible=True, disabled=True, size=(80, 1)),
     gui.FilesBrowse(target='PathInput',
                     visible=True, key='Browse', disabled=True)],
    [gui.Text(text='', key='Description')],
    [Collapsible(Options, 'Options', 'Options', collapsed=True)],
    [gui.Submit(button_text='Run', size=15, visible=False, key='Run')],
    [gui.ProgressBar(max_value=100, orientation='horizontal', size=(50, 20),
     bar_color=('green', 'white'), key='PBar', visible=False)],
    [gui.Text(text='', key='PBarFile')],
    [gui.Text(text='', key='PBarFileStep')]
]

layout = [[TopText, gui.Column(Sidebar, vertical_scroll_only=True,
                               scrollable=True, expand_y=True),
          gui.VSeparator(),
          gui.Column(Rightlayout)]]

MainWindow = gui.Window('Prep ToolKit', layout, size=(850, 400))
Function = ''

while True:
    event, values = MainWindow.read()
    match event:
        case 'Exit' | gui.WIN_CLOSED:
            break
        case 'Run':
            PathList = values['PathInput'].split(';')
            MainWindow['PBar'].update(visible=True)

            for index, PathInput in enumerate(PathList):
                if not isfile(PathInput):
                    gui.popup_error(f'the following file:\n{PathInput}\n\
                    is not a valid file!\n Skipping this file',
                                    title='Invalid file error',
                                    auto_close=True, auto_close_duration=5,
                                    keep_on_top=True, modal=True)
                    continue
                if Break is True:
                    Break = False
                    break
                match Function:
                    case 'Contact_Sheet':
                        Contact_Sheet(PathInput)
                    case 'Bilingual_Table':
                        Bilingual(PathInput)
                    case 'Doc2PDF':
                        break
                    case 'Accept Revisions':
                        break
                    case 'Prep_Story':
                        PrepStoryExport(PathInput)
                    case 'Unhide':
                        Unhide(PathInput)
            MainWindow['PBar'].update(100)
            MainWindow['PBarFile'].update('')
            MainWindow['PBarFileStep'].update('Done!')
        case 'Contact Sheet':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_types['cs'],
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Contact_Sheet')
        case 'Options-BUTTON-':
            MainWindow['Options'].update(visible=not MainWindow['Options'].
                                         visible)
            MainWindow['Options'+'-BUTTON-'].\
                update(MainWindow['Options'].metadata[0] if
                       MainWindow['Options'].visible else
                       MainWindow['Options'].metadata[1])
        case 'Bilingual Table':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_types['bil'],
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Bilingual_Table')
        case 'Word to PDF':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_types['pdf'],
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Doc2PDF')
        case 'Accept Revisions':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_types['rev'],
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Accept_Revisions')
        case 'Prep Story Export':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_types['story'],
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Prep_Story')
        case 'Unhide':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_types['unhide'],
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Unhide')
MainWindow.close()
