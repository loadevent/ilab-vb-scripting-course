'********************************************************
'          Tutorial 9 - Read From Excel
'********************************************************
REM This script read data from excel spreadsheet

'Declare
    DIM objExcel
    DIM objWorkbook
    DIM strFilePath
    DIM strOutput

'Assign
    strFilePath = "C:\Users\Kabelo Tlhape\Music\excel_read.xls"
    SET objExcel = CreateObject("Excel.Application")
    SET objWorkbook = objExcel.Workbooks.Open(strFilePath)

'Use

    DIM intRow
    intRow = 2

    strOutput = strOutput & objExcel.Cells(1,1) & "   " & objExcel.Cells(1,2).Value & vbNewLine _
    & "========================" & vbNewLine

    DO UNTIL objExcel.Cells(intRow, 1).Value = ""
        strOutput = strOutput & objExcel.Cells(intRow, 1).Value & "            " & objExcel.Cells(intRow,2).Value & vbNewLine
        intRow = intRow +1
    LOOP
    MsgBox strOutput

    objExcel.Quit