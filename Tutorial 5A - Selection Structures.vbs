'********************************************************
'            Tutorial 5 - Selection Structures
'********************************************************
REM This tutorial teaches the different selection Structures

'Declare
    DIM intAge
    DIM strResults

'Assign
    intAge = 2

    'unary if
    IF intAge >= 18 THEN
        strResults = "Adult"
    END IF


    'binary IF
    IF intAge >= 18 THEN
        strResults = "Adult"
    ELSE
        strResults = "Not an adult"
    END IF

     MsgBox("Age: " & intAge & vbNewLine _
                    & "Results: " & strResults)

    IF intAge >= 18 THEN
        strResults = "Adult"
    ELSEIF intAge >= 13 THEN
        strResults = "Teenager"
    ELSEIF intAge >= 0 THEN
        strResults = "Kid"
    ELSE
        strResults = "Age cannot be negative"
    END IF

'Use
    MsgBox("Age: " & intAge & vbNewLine _
                    & "Results: " & strResults)
