'create a script that will write the following data to excel
'workbook:
'student name, test mark, assignment mark, final mark
'final mark = test mark + assignment divide by 2
'write at least 3 entries

Option Explicit

'Declare
DIM objExcel
DIM strFilePath
DIM objFso
DIM objWorkbook
DIM numOfStudents
DIM strStName, intTest, intAssignment, dblFinalMark
DIM i

strFilePath = "C:\Users\Kabelo Tlhape\Music\excel_write.xls"
    DO 
        numOfStudents = CInt(InputBox("How many student would you like to capture?"))
        IF numOfStudents < 3 THEN
            MsgBox("At least 3 students")
        END IF
    LOOP while numOfStudents < 3

    SET objExcel = CreateObject("Excel.Application")
    SET objFso = CreateObject("Scripting.FileSystemObject")
    SET objWorkbook = objExcel.Workbooks.Open(strFilePath)

    objExcel.Workbooks.Add
    objExcel.Worksheets(1).Cells(1,1).Value = "Student Name"
    objExcel.Worksheets(1).Cells(1,2).Value = "Test Mark"
    objExcel.Worksheets(1).Cells(1,3).Value = "Assignment Mark"
    objExcel.Worksheets(1).Cells(1,4).Value = "Final Mark"

    FOR i = 2 to (numOfStudents + 1)
        strStName = InputBox("Student Name")
        intTest = CInt(InputBox("Test Mark"))
        intAssignment = CInt(InputBox("Assignment Mark"))
        dblFinalMark = (intTest + intAssignment) / 2

        objExcel.Worksheets(1).Cells(i,1).Value = strStName
        objExcel.Worksheets(1).Cells(i,2).Value = intTest
        objExcel.Worksheets(1).Cells(i,3).Value = intAssignment
        objExcel.Worksheets(1).Cells(i,4).Value = dblFinalMark
    NEXT
    objExcel.Quit

'WITH objExcel.Worksheets(1)
   '.Cells(1,1).Value = "Student Name"
   '.Cells(1,2).Value = "Test Mark"
   '.Cells(1,3).Value = "Assignment Mark"
   '.Cells(1,4).Value = "Final Mark"
   '.Cells(2,1).Value = "Luvuyo"
   '.Cells(2,2).Value = 100
   '.Cells(2,3).Value = 100
   '.Cells(2,4).Formula = "=(B2 + C2) / 2"

   '.Cells(3,1).Value = "Dave"
   '.Cells(3,2).Value = 100
   '.Cells(3,3).Value = 100
   '.Cells(3,4).Formula = "=(B3 + C3) / 2"

   '.Cells(4,1).Value = "Thato"
   '.Cells(4,2).Value = 100
   '.Cells(4,3).Value = 100
   '.Cells(4,4).Formula = "=(B4 + C4) / 2"

'END WITH

'objWorkbook.SaveAs(strFilePath)
'objWorkbook.close
'objExcel.Quit