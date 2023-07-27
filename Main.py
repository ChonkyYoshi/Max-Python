import PySimpleGUI as gui
import Functions as fn
from configparser import ConfigParser
import win32com.client as com
from pathlib import Path

Break = False
Upsaved = False
config = ConfigParser()
config.read('config.ini')
file_ext = dict(config['file_ext'])
for key in file_ext:
    file_ext[key] = tuple(file_ext[key].split(','))  # type: ignore


def ClearOptions():
    for element in MainWindow.element_list():
        if element.metadata == 'option':
            element.update(text='')
            element.update(visible=False)
    MainWindow['UserRegex'].update(visible=False)


def SetOptions(Function):
    match Function:
        case 'Contact_Sheet':
            ClearOptions()
            MainWindow['R1O1'].update(text='Compress images')
            MainWindow['R1O1'].update(visible=True)
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(value=config['Descriptions']
                                             ['Contact_Sheet'])
        case 'Bilingual_Table':
            ClearOptions()
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(value=config['Descriptions']
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
            MainWindow['Description'].update(value=config['Descriptions']
                                             ['Doc2PDF'])
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
            MainWindow['Description'].update(value=config['Descriptions']
                                             ['Accept_Revisions'])
        case 'Prep_Story':
            ClearOptions()
            MainWindow['UserRegex'].update(visible=True)
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(value=config['Descriptions']
                                             ['Prep_Story'])
        case 'Unhide':
            ClearOptions()
            MainWindow['R1O1'].update(text='Excel: Skip hidden rows')
            MainWindow['R1O1'].update(visible=True)
            MainWindow['R2O1'].update(text='Excel: Skip hidden columns')
            MainWindow['R2O1'].update(visible=True)
            MainWindow['R3O1'].update(text='Excel: Skip hidden sheets')
            MainWindow['R3O1'].update(visible=True)
            MainWindow['R1O2'].update(text='Powerpoint: Skip hidden shapes')
            MainWindow['R1O2'].update(visible=True)
            MainWindow['R2O2'].update(text='Powerpoint: Skip hidden slides')
            MainWindow['R2O2'].update(visible=True)
            MainWindow['R3O2'].update(text='Global: Overwrite')
            MainWindow['R3O2'].update(visible=True)
            MainWindow['Run'].update(visible=True)
            MainWindow['Description'].update(value=config['Descriptions']
                                             ['Unhide'])
    return Function


def Collapsible(layout, key, title='', arrows=(gui.SYMBOL_DOWN, gui.SYMBOL_UP),
                collapsed=False):
    return gui.Column([[gui.T((arrows[1] if collapsed else arrows[0]),
                      enable_events=True, k=key+'-BUTTON-'), gui.T(title,
                      enable_events=True, key=key+'-TITLE-')],
                      [gui.pin(gui.Column(layout, key=key,
                       visible=not collapsed, metadata=arrows))]], pad=(0, 0))


def Contact_Sheet(File: Path):
    if File.suffix in ['.doc', '.ppt', '.xls']:
        Upsaved = True
        MainWindow['PBarFile'].update(value=File.name)
        MainWindow['PBarFileStep'].update(value='Upsaving to Office 2007 ' +
                                          'format')
        MainWindow['PBar'].update(current_count=(index+1/5)/len(PathList)*100)
        File = fn.Upsave(File)
    MainWindow['PBarFile'].update(value=File.name)
    MainWindow['PBarFileStep'].update(value='Extracting Images')
    MainWindow['PBar'].update(current_count=(index+2/5)/len(PathList)*100)
    TempDir = fn.ExtractImages(File)
    for step in fn.CleanTempDir(TempDir,
                                MainWindow['R1O1'].get()):  # type: ignore
        MainWindow['PBarFileStep'].update(value=step)
        MainWindow['PBar'].update(current_count=(index+3/5)/len(PathList)*100)
    for step in fn.FillCS(TempDir, File):
        MainWindow['PBarFileStep'].update(value=step)
        MainWindow['PBar'].update(current_count=(index+4/5)/len(PathList)*100)
        MainWindow.refresh()
    if Upsaved:
        File.unlink()
        Upsaved = False


def Bilingual(File: Path):
    if File.suffix in ['.doc', '.ppt', '.xls']:
        Upsaved = True
        MainWindow['PBarFile'].update(value=File.name)
        MainWindow['PBarFileStep'].update(value='Upsaving to Office 2007 ' +
                                          'format')
        MainWindow['PBar'].update(current_count=(index+1/5)/len(PathList)*100)
        File = fn.Upsave(File)
    MainWindow['PBarFile'].update(text=File.name)
    for step in (fn.BilTable(File)):
        MainWindow['PBar'].update(current_count=(index+2/3)/len(PathList)*100)
        MainWindow['PBarFileStep'].update(text=step)
        MainWindow.refresh()
    if Upsaved:
        PathInput.unlink()
        Upsaved = False


def Doc2PDF(WordApp, File: Path):
    if MainWindow['R1O1'].get() is True and MainWindow['R2O1'].get() is True:
        gui.popup_error('Both Accept and Reject revisions are ticked!\n\
        Please choose only one and try again', title='impossible options',
                        modal=True)
        global Break
        Break = True
    else:
        MainWindow['PBarFile'].update(text=File.name)
        MainWindow['PBarFileStep'].update(text='Saving as PDF...')
        MainWindow['PBar'].update(current_count=index/len(PathList)*100)
        fn.Doc2PDF(WordApp, File,
                   ARev=MainWindow['R1O1'].get(),  # type: ignore
                   DRev=MainWindow['R2O1'].get(),  # type: ignore
                   Com=MainWindow['R3O1'].get(),  # type: ignore
                   Overwrite=MainWindow['R2O1'].get())  # type: ignore
    if index + 1 == len(PathList):
        WordApp.Quit()


def AcceptRevisions(WordApp, File: Path):
    if MainWindow['R1O1'].get() is True and MainWindow['R2O1'].get() is True:
        gui.popup_error('Both Accept and Reject revisions are ticked!\n\
        Please choose only one and try again', title='impossible options',
                        modal=True)
        global Break
        Break = True
    else:
        MainWindow['PBarFile'].update(text=File.name)
        MainWindow['PBarFileStep'].update(text='Accepting revisions...')
        MainWindow['PBar'].update(text=index/len(PathList)*100)
        fn.AcceptRevisions(WordApp, File,
                           ARev=MainWindow['R1O1'].get(),  # type: ignore
                           DRev=MainWindow['R2O1'].get(),  # type: ignore
                           Com=MainWindow['R3O1'].get(),  # type: ignore
                           Overwrite=MainWindow['R2O1'].get())  # type: ignore
    if index + 1 == len(PathList):
        WordApp.Quit()


def PrepStoryExport(File: Path):
    MainWindow['PBarFile'].update(text=File)
    for step in fn.PrepStoryExport(File, Regex):
        MainWindow['PBar'].update(current_count=index/len(PathList)*100)
        MainWindow['PBarFileStep'].update(text=str(step))
        MainWindow.refresh()


def Unhide(File: Path):
    if File.suffix in ['.doc', '.ppt', '.xls']:
        Upsaved = True
        MainWindow['PBarFile'].update(value=File.name)
        MainWindow['PBarFileStep'].update(value='Upsaving to Office 2007 ' +
                                          'format')
        MainWindow['PBar'].update(current_count=(index+1/3)/len(PathList)*100)
        File = fn.Upsave(File)
    MainWindow['PBarFile'].update(text=File.name)
    for step in fn.Unhide(File,
                          Row=MainWindow['R1O1'].get(),  # type: ignore
                          Col=MainWindow['R2O1'].get(),  # type: ignore
                          Sheet=MainWindow['R3O1'].get(),  # type: ignore
                          Shp=MainWindow['R1O2'].get(),  # type: ignore
                          Sld=MainWindow['R2O2'].get(),  # type: ignore
                          Overwrite=MainWindow['R3O2'].get()):  # type: ignore
        MainWindow['PBar'].update(current_count=(index + 2/3) /
                                  len(PathList)*100)
        MainWindow['PBarFileStep'].update(text=str(step))
        MainWindow.refresh()
    if Upsaved:
        File.unlink()
        Upsaved = False


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
     key='R3O4', visible=False, size=(22, 1), pad=(0, 0), metadata='option')],
    [gui.Input(default_text='', key='UserRegex', visible=False,
               do_not_clear=False)]
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

MainWindow = gui.Window('Prep ToolKit', layout, size=(850, 500))
Function = ''

while True:
    event, values = MainWindow.read()
    match event:
        case 'Exit' | gui.WIN_CLOSED:
            break
        case 'Run':
            PathList = values['PathInput'].split(';')
            MainWindow['PBar'].update(visible=True)
            Regex = values['UserRegex']
            for index, PathInput in enumerate(PathList):
                PathInput = Path(PathInput)
                if not PathInput.is_file:
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
                        try:
                            WordApp = com.GetActiveObject('Word.Application')
                        except Exception:
                            WordApp = com.DispatchEx('Word.Application')
                        Doc2PDF(WordApp, PathInput)
                    case 'Accept Revisions':
                        try:
                            WordApp = com.GetActiveObject('Word.Application')
                        except Exception:
                            WordApp = com.DispatchEx('Word.Application')
                            AcceptRevisions(WordApp, PathInput)
                    case 'Prep_Story':
                        PrepStoryExport(PathInput)
                    case 'Unhide':
                        Unhide(PathInput)
            MainWindow['PBar'].update(current_count=100)
            MainWindow['PBarFile'].update(value='')
            MainWindow['PBarFileStep'].update(text='Done!')
        case 'Contact Sheet':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_ext['cs'],  # type: ignore
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
            MainWindow['Browse'].FileTypes = file_ext['bil'],  # type: ignore
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Bilingual_Table')
        case 'Word to PDF':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_ext['pdf'],  # type: ignore
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Doc2PDF')
        case 'Accept Revisions':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_ext['rev'],  # type: ignore
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Accept_Revisions')
        case 'Prep Story Export':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_ext['story'],  # type: ignore
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Prep_Story')
        case 'Unhide':
            MainWindow['Browse'].update(disabled=False)
            MainWindow['Browse'].FileTypes = file_ext['unh'],  # type: ignore
            MainWindow['PathInput'].update(disabled=False)
            Function = SetOptions('Unhide')
MainWindow.close()
