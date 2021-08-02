import openpyxl
import glob
import os

fileDir  = r"D:/tmp/verifi/"
fileExet = ".xlsx"
TARGET_ZOOM_SCALE = 85

def checkProperties(workbook):
    print('***PROPERTIES CHECK***')
    if workbook.properties.creator is None:
        print('OK\n')
    else:
        print('NG')
        print('creator is ' + workbook.properties.creator + '\n')

def checkZoomScale(sheet):
    print('+- ZOOM SCALE CHECK -+')
    if sheet.sheet_view.zoomScale == TARGET_ZOOM_SCALE:
        """
        print('Zoomscale:%3d' % sheet.sheet_view.zoomScale)
        print('Audit passed\n')
        """
        return True
    elif sheet.sheet_view.zoomScale is None:
        print('Zoomscale:100')
    else:
        print('Zoomscale:%3d' % sheet.sheet_view.zoomScale)
    """
    print('Audit NOT passed\n')
    """
    return False

def checkActiveCell(sheet):
    print('+- ACTIVE CELL CHECK -+')
    for idx in range(len(sheet.sheet_view.selection)):
        if sheet.sheet_view.selection[idx].activeCell == 'A1':
            print('ActiveCell[%s] Selection[%s]' % (sheet.sheet_view.selection[idx].activeCell, idx))
            """
            print('Audit passed\n')
            """
            return True
        else:
            print('ActiveCell[%s] Selection[%s]' % (sheet.sheet_view.selection[idx].activeCell, idx))

    """
    print('Audit NOT passed')
    """
    return False

for name in glob.glob(fileDir + '**/*' + fileExet , recursive=True):
    print('+------ File[%s] ------+' % name)
    wb = openpyxl.load_workbook(name,read_only=False)

    checkProperties(wb)

    for sheet in wb:
        print('+--- Sheet Name[%s] ---------+' % sheet.title)
        print('OK') if checkZoomScale(sheet) == True else print('NG')
        print('OK') if checkActiveCell(sheet) == True else print('NG')
        print('')
