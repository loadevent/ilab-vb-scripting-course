REM Using nested conditional statements and logical operators

    DIM intNumber
    DIM strResults

    intNumber = 38

    '25 OR 40
    IF intNumber > 20 THEN
        'TRUE
        IF intNumber = 25 THEN
            strResults = "The number is: " & intNumber
        ELSEIF intNumber = 40 THEN
            strResults = "The number is: " & intNumber
        ELSE
            strResults = "Not the number we are looking for: [" & intNumber & "]"
        END IF
    ELSE
        strResults = "Number [" & intNumber & "] is not greater than 20"
    END IF

    MsgBox(strResults)