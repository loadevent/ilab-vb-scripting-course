REM This tutorial teaches you about different iteration Structures
'********************************************************
'                  Tutorial 6 - Looping
'********************************************************

'display Tuesday 3 times using a FOR loop

    'FOR x = 1 to 3
    '    MsgBox(x & " - Tuesday")
    'NEXT

'display even numbers between 1 and 10 using FOR Loop
    'FOR x = 2 to 8 STEP 2
    '    MsgBox(x)
    'NEXT

'display Monday 3 times using while loop
    DIM intCount

    intCount = 1'...7

    DO WHILE intCount <= 3
        'MsgBox(intCount & " - Monday")
        intCount = intCount + 1
    LOOP

'display Thursday 4 times using Do Until
    DIM intCountUntil

    intCountUntil = 1

    DO UNTIL intCountUntil = 4
        'MsgBox("Thursday")
        intCountUntil = intCountUntil + 1
    LOOP

'display Friday 3 times using a DO loop

    intCount = 1

    DO
        MsgBox("Friday")'1 2
        intCount = intCount + 1
    LOOP UNTIL intCount = 4


