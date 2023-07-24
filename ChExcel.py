import openpyxl as xl
from openpyxl.utils.cell import get_column_letter
# from json import dumps
import regex as reg


def GetDefinedNameData(defined_name, dict):
    dict['Name'] = defined_name.name
    dict['Type'] = defined_name.type
    dict['Value'] = defined_name.value


def GetCommentData(Comment, dict):
    dict['Content'] = Comment.content
    dict['Author'] = Comment.author


def GetHyperlinkData(cell, dict):
    dict['ref'] = cell.hyperlink.ref
    dict['Location'] = cell.hyperlink.location
    dict['Tooltip'] = cell.hyperlink.tooltip
    dict['display'] = cell.hyperlink.display
    dict['Target'] = cell.hyperlink.target
    dict['Text'] = cell.internal_value


def GetHiddenRows(Sheet):
    hidden_rows = list()
    for row in range(Sheet.min_row, Sheet.max_row + 1):
        if Sheet.row_dimensions[row].hidden:
            if row not in hidden_rows:
                hidden_rows.append(row)
    return hidden_rows


def GetHiddenColumns(Sheet):
    hidden_columns = list()
    for col in range(Sheet.min_column, Sheet.max_column + 1):
        if Sheet.column_dimensions[get_column_letter(col)].hidden:
            if col not in hidden_columns:
                hidden_columns.append(col)
                if Sheet.column_dimensions[get_column_letter(col)]\
                        .outlineLevel == 1:
                    for i in range(Sheet.column_dimensions[get_column_letter
                                                           (col)].min,
                                   Sheet.column_dimensions[get_column_letter
                                                           (col)].max):
                        if i not in hidden_columns:
                            hidden_columns.append(i)
    return hidden_columns


def GetTableData(table, dict):
    dict['Name'] = table.name
    dict['Display Name'] = table.displayName
    dict['Range'] = table.ref


def GetDataValidationData(datavalidation):
    Infos = dict()
    Infos['Prompt'] = datavalidation.prompt
    Infos['Prompt Title'] = datavalidation.promptTitle
    Infos['Error message'] = datavalidation.error
    Infos['Error Style'] = datavalidation.errorStyle
    Infos['Error Title'] = datavalidation.errorTitle
    Infos['Formula'] = datavalidation.formula1
    Infos['Type'] = datavalidation.type
    Infos['Range'] = list()
    for range in datavalidation.cells.ranges:
        Infos['Range'].append(range.coord)
    return Infos


Results = dict()
wb = xl.open('C:/Users/transformers/Downloads/FlagsAll.xlsm', keep_vba=True)
Results['Global defined names'] = dict()
for name in wb.defined_names:
    Results['Global defined names'][name] = dict()
    GetDefinedNameData(wb.defined_names[name],
                       Results['Global defined names'][name])
Results['Sheets'] = dict()
for index, ws in enumerate(wb.worksheets):
    Results['Sheets'][ws.title] = dict(Visibility=ws.sheet_state)
    if len(ws.defined_names) > 0:
        Results['Sheets'][ws.title]['Sheet defined names'] = dict()
        for name in ws.defined_names:
            Results['Sheets'][ws.title]['Sheet defined names'][name] = dict()
            GetDefinedNameData(ws.defined_names[name],
                               Results['Sheets'][ws.title]
                               ['Sheet defined names'][name])
    if len(ws.tables) > 0:
        Results['Sheets'][ws.title]['Sheet Tables'] = dict()
        for table in ws.tables:
            Results['Sheets'][ws.title]['Sheet Tables'][table] = dict()
            GetTableData(ws.tables[table],
                         Results['Sheets'][ws.title]['Sheet Tables'][table])
            if ws.protection.sheet is True:
                Results['Sheets'][ws.title]['Protection'] = 'Protected'
            else:
                Results['Sheets'][ws.title]['Protection'] = 'Unlocked'
    Results['Sheets'][ws.title]['Hidden rows'] = GetHiddenRows(ws)
    Results['Sheets'][ws.title]['Hidden columns'] = GetHiddenColumns(ws)
    if len(ws.data_validations.dataValidation) > 0:
        Results['Sheets'][ws.title]['Data Validation'] = list()
        for data in ws.data_validations.dataValidation:
            Results['Sheets'][ws.title]['Data Validation'].append(
                GetDataValidationData(data))
    Results['Sheets'][ws.title]['Cells'] = dict()
    for row in range(ws.min_row, ws.max_row + 1):
        for col in range(ws.min_column, ws.max_column + 1):
            cell = ws.cell(row, col)
            Results['Sheets'][ws.title]['Cells'][cell.coordinate] = dict()
            celldict = Results['Sheets'][ws.title]['Cells'][cell.coordinate]
            if cell.comment is not None:
                (celldict['Comment']) = dict()
                GetCommentData(cell.comment, celldict['Comment'])
            if cell.data_type == 'f':
                if cell.internal_value.startswith('='):
                    (celldict['Formula']) = str(cell.internal_value)
            if reg.search(r'<[^>]+>', str(cell.internal_value)):
                celldict['html'] = True
            if reg.search(r'{.*?}', str(cell.internal_value)):
                celldict['Variables'] = True
            if cell.hyperlink != '' and cell.hyperlink is not None:
                (celldict['Hyperlink']) = dict()
                GetHyperlinkData(cell, celldict['Hyperlink'])
            if len(str(cell.internal_value)) > 25000:
                celldict['Long Text'] = True
            if not celldict:
                Results['Sheets'][ws.title]['Cells'].pop(cell.coordinate)
# open('results.json', 'w').write(dumps(Results))
