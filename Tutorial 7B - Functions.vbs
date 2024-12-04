'********************************************************
'                   Tutorial 7 - Functions
'********************************************************
REM This script theaches you about Functions

Function getFullname()
    DIM strFirstname, strLastname, strFullname
    'Assign
    strFirstname = Trim(InputBox("Please enter Firstname"))
    strLastname =  Trim(InputBox("Please enter Lastname"))
    strFullname = strFirstname & " " & strLastname
    'return the fullname (strFullname)
    getFullname = strFullname

end function

'================================================================
'Declare
    DIM strFullName

'Assign
    strFullName = getFullname()
'Use
    MsgBox("Fullname: " & strFullName)
