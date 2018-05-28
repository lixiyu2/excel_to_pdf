# -*- coding: utf-8 -*-
from win32com import client
import os
import sys


def main(args):
    input_dir = args[0]
    out_dir = args[1]
    for root, dirs, files in os.walk(input_dir):
        for name in files:
            fName = os.path.join(root, name)
            name = name.split(".")[0]
            # arguments for excel file and output path
            excelFile = os.path.normpath(fName)
            pdfPath = os.path.normpath(out_dir)
            # create Excel object and open workbook
            objExcel = client.DispatchEx("Excel.Application")
            objExcel.Visible = False
            workbook = objExcel.Workbooks.Open(excelFile, 1)
            workbook.saved = True

            # refresh workbook
            workbook.RefreshAll()
            # loop through sheets, ignoring any specified in ignoreSheets
            for index in range(workbook.Sheets.Count):
                # if index not in ignoreSheets:
                sheet = workbook.Worksheets[index]
                # create path for pdf
                filePath = os.path.join(pdfPath, '%s.pdf' % (name))
                # ignore empty worksheets
                if sheet.UsedRange.Columns.Count == 1 & sheet.UsedRange.Rows.Count == 1:
                    print('sheet %s is empty; could not save to pdf' % (name))
                # export as pdf
                else:
                    try:
                        sheet.ExportAsFixedFormat(0, filePath)
                        print('saving %s.pdf' % (name))
                    except Exception, e:
                        print e
                        print('could not convert sheet "%s"' % (name))

            # close Excel
            objExcel.Workbooks.Close()
            objExcel.Quit()


if __name__ == '__main__':
    main(sys.argv[1:])
