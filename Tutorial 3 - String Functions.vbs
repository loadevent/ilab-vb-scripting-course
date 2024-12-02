REM This tutorial teaches you about string functions
'********************************************************
'             Tutorial 3 - String Functions
'********************************************************

'Declare
    DIM strFirstname, strLastname
    DIM strFullname
    DIM intStringPos
    DIM strSearchString
    DIM strCourseName, strCourseInput
    DIM strStringCompare

'Assign
    strFirstname = Trim(UCase(InputBox("Please enter first name")))
    strLastname = Trim(UCase(InputBox("Please enter last name")))

    strFullname = strFirstname & " " & strLastname
    strCourseName = "VB Scripting"


'Use / Consume
    MsgBox("Firstname: " & strFirstname & vbNewLine _
                        & "Lastname: " & strLastname & vbNewLine _
                        & "================================" & vbNewLine _
                        & "Fullname: " & strFullname & vbNewLine _
                        & strFirstname & " has " & Len(strFirstname) & " characters" & vbNewLine _
                        & "The first 3 characters of " & strFirstname & " are: " & Left(strFirstname, 3))
    'Capture the string to search for
    strSearchString = InputBox("Type the letter to search from the name (" & strFirstname & ")")
    intStringPos = InStr(1, strFirstname, strSearchString)

    MsgBox(strSearchString & " is found at position: " & intStringPos)

    strCourseInput = InputBox("Guess the name of the course!!!")
    strStringCompare = StrComp(strCourseInput, strCourseName)
    MsgBox("Comparison between " & strCourseInput & " and " & strCourseName & " is: " & strStringCompare)




