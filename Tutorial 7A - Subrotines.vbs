'********************************************************
'                   Tutorial 7 - Subroutines
'********************************************************
REM This script theaches you about subroutines

    'create a subroutine that will display a welcome message
    Sub showMessage(strName)
        MsgBox("Hi " & strName & ", welcome to VB Scripting Course")
    end sub

    sub showFullname(strFirstname, strLastname)
        DIM strFullname

        strFullname = strFirstname & " " & strLastname
        MsgBox("Fullname: " & strFullname)
    end sub
'============================================================================
    'declare
    DIM strName, strSurname

    strName = InputBox("Please enter your name")
    strSurname = InputBox("Please enter your surname")
    Call showMessage(strName)
    Call showFullname(strName, strSurname)