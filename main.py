import openpyxl
import glob
import os

fileDir  = r"D:/tmp/test/"
fileExet = ".xlsx"

for name in glob.glob(fileDir + '**/*' + fileExet , recursive=True):
    print(name)
    wb = openpyxl.load_workbook(name,read_only=False)

    # check properties data
    if wb.properties.creator is None:
        print('PROPERTIES CHECK OK')
    else:
        print('PROPERTIES CHECK NG')
        print('creator:' + wb.properties.creator)

    # check Zoomscale and active cell
    for sheet in wb:
        if sheet.sheet_view.zoomScale is None:
            print(sheet.title + '---' + 'Zoomscale:100' + '   Active cell:' + sheet.sheet_view.selection[0].activeCell)
        else:
            print(sheet.title + '---' + 'Zoomscale:%3d' % sheet.sheet_view.zoomScale + '   Active cell:' + sheet.sheet_view.selection[0].activeCell)
    
    print('')
