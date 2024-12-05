'********************************************************
'          Tutorial 9 - Write To Excel
'********************************************************
REM This script will write data to excel workbook
Option Explicit

'Declare
    DIM objExcel 'object used to open excel workbook
    DIM strFilePath 'location of the workbook
    DIM objFileSystemObject

'Assign
    SET objExcel = CreateObject("Excel.Application")
    SET objFileSystemObject = CreateObject("Scripting.FileSystemObject")
    strFilePath = "C:\Users\Kabelo Tlhape\Music\excel_write.xls"

    IF objFileSystemObject.FileExists(strFilePath) THEN
        objExcel.Workbooks.Open(strFilePath) 'open a workbook
        objExcel.ActiveWorkbook.Save

    ELSE
        objExcel.Workbooks.Add
        objExcel.Worksheets(1).Cells(1,1).Value = "City"
        objExcel.Worksheets(1).Cells(1,2) = "Population"
        objExcel.Worksheets(1).Cells(2,1) = "Pretoria"
        objExcel.Worksheets(1).Cells(2,2) = 900000
        objExcel.Worksheets(1).Cells(3,1) = "Midrand"
        objExcel.Worksheets(1).Cells(3,2) = 100000
        objExcel.Worksheets(1).Cells(4,1) = "Jo'burg"
        objExcel.Worksheets(1).Cells(4,2) = 120000
        objExcel.ActiveWorkbook.SaveAs(strFilePath)

    END IF
   
    objExcel.Quit 'close or dispose the object

    







