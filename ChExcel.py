import openpyxl as xl
from openpyxl.utils.cell import get_column_letter
from json import dumps

Results = dict()

wb = xl.open('C:/Users/transformers/Downloads/FlagsAll.xlsm', keep_vba=True)
Results['Global defined names'] = dict()
for name in wb.defined_names:
    Results['Global defined names'][name] = dict()
    Results['Global defined names'][name]['Name'] = wb.defined_names[name].name
    Results['Global defined names'][name]['Type'] = wb.defined_names[name].type
    Results['Global defined names'][name]['Value'] = (wb.defined_names[name].
                                                      value)
Results['Sheets'] = dict()
for index, ws in enumerate(wb.worksheets):
    Results['Sheets'][ws.title] = dict()
    Results['Sheets'][ws.title]['Visibility'] = ws.sheet_state
    Results['Sheets'][ws.title]['Sheet defined names'] = dict()
    for name in ws.defined_names:
        Results['Sheets'][ws.title]['Sheet defined names'][name] = dict()
        Results['Sheets'][ws.title]['Sheet defined names'][name]['Name'] =\
            ws.defined_names[name].name
        Results['Sheets'][ws.title]['Sheet defined names'][name]['Type'] =\
            ws.defined_names[name].type
        Results['Sheets'][ws.title]['Sheet defined names'][name]['Value'] =\
            ws.defined_names[name].value
    if ws.protection.sheet is True:
        Results['Sheets'][ws.title]['Protection'] = 'Protected'
    else:
        Results['Sheets'][ws.title]['Protection'] = 'Unlocked'
    hidden_rows = list()
    hidden_columns = list()
    for row in range(ws.min_row, ws.max_row + 1):
        if ws.row_dimensions[row].hidden:
            if row not in hidden_rows:
                hidden_rows.append(row)
    Results['Sheets'][ws.title]['Hidden rows'] = hidden_rows
    for col in range(ws.min_column, ws.max_column + 1):
        if ws.column_dimensions[get_column_letter(col)].hidden:
            if col not in hidden_columns:
                hidden_columns.append(col)
            if ws.column_dimensions[get_column_letter(col)].outlineLevel == 1:
                for i in range(ws.column_dimensions[get_column_letter(col)].
                               min,
                               ws.column_dimensions[get_column_letter(col)].
                               max):
                    if i not in hidden_columns:
                        hidden_columns.append(i)
    Results['Sheets'][ws.title]['Hidden columns'] = hidden_columns
    Results['Sheets'][ws.title]['Cells'] = dict()
    for row in range(ws.min_row, ws.max_row + 1):
        for col in range(ws.min_column, ws.max_column + 1):
            cell = ws.cell(row, col)
            Results['Sheets'][ws.title]['Cells'][cell.coordinate] = dict()
            if cell.comment is not None:
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Comment']) = dict()
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Comment']['Content']) = cell.comment.content
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Comment']['Author']) = cell.comment.author
            if cell.data_type == 'f' and cell.internal_value.startswith('='):
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Formula']) = str(cell.internal_value)
            if cell.hyperlink != '' and cell.hyperlink is not None:
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Hyperlink']) = dict()
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Hyperlink']['ref']) = cell.hyperlink.ref
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Hyperlink']['location']) = cell.hyperlink.location
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Hyperlink']['tooltip']) = cell.hyperlink.tooltip
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Hyperlink']['display']) = cell.hyperlink.display
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Hyperlink']['Target']) = cell.hyperlink.target
                (Results['Sheets'][ws.title]['Cells'][cell.coordinate]
                 ['Hyperlink']['Text']) = cell.internal_value
            if not Results['Sheets'][ws.title]['Cells'][cell.coordinate]:
                Results['Sheets'][ws.title]['Cells'].pop(cell.coordinate)
open('results.json', 'w').write(dumps(Results))
