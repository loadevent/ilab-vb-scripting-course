'Create a program that will prompt the user for any integer 
'except for zero, the program should add the integers together
'and display their sum. The program should stop when the user 
'enters zero
    DIM intNumber, intTotal

    intTotal = 0
    intNumber = CInt(InputBox("Enter any number"))

    DO UNTIL intNumber = 0
        intTotal = intTotal + intNumber
        intNumber = CInt(InputBox("Enter any number"))
    LOOP

    'intNumber = CInt(InputBox("Enter any number"))
'
    'DO WHILE intNumber <> 0
    '    intTotal = intTotal + intNumber
    '    intNumber = CInt(InputBox("Enter any number"))
    'LOOP

    'DO 
    '    intNumber = CInt(InputBox("Enter any number"))
    '    IF intNumber <> 0 THEN
    '        intTotal = intTotal + intNumber
    '    END IF
    'LOOP UNTIL intNumber = 0

'DO 
'    intNumber = InputBox("Enter any number")
'    IF IsNumeric(intNumber) THEN
'        IF intNumber = 0 THEN
'            EXIT DO
'        END IF
'    ELSE 
'        MsgBox("Please enter numeric value")
'    END IF
'    intTotal = intTotal + intNumber
'LOOP 

MsgBox("Sum of intergers: " & intTotal)
