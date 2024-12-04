'display the following menu to the user:

 'Please enter any of these operators:
 '+ Addition
 '- Subtraction
 '* Multiply
 '/ Divide
 
 'the program should allow the user to use one operator
 'from the menu. Display the menu as long as the user
 'enters an incorrect operator
    DIM strOperator, blIsValid

    MsgBox("Please enter any of these operators" & vbNewLine _
                                & "+ Addition" & vbNewLine _
                                & "- Subtraction" & vbNewLine _
                                & "* Multiply" & vbNewLine _
                                & "/ Divide")
    strOperator = Trim(InputBox("Your selection"))

    DO WHILE strOperator <> "+" AND strOperator <> "-" AND strOperator <> "*" AND strOperator <> "/" 
          strOperator = Trim(InputBox("Your selection"))
    LOOP



    'DO 
    '    strOperator = InputBox("Please enter any of these operators" & vbNewLine _
    '                            & "+ Addition" & vbNewLine _
    '                            & "- Subtraction" & vbNewLine _
    '                            & "* Multiply" & vbNewLine _
    '                            & "/ Divide")
    '    IF strOperator = "+" OR strOperator = "-" OR strOperator = "*" OR strOperator = "/" THEN
    '        EXIT DO
    '    ELSE
    '        MsgBox("Invalid Operator, please try again")
    '    END IF
    'LOOP

    'blIsValid = FALSE

    'DO UNTIL blIsValid = TRUE
    '    strOperator = InputBox("Please enter any of these operators" & vbNewLine _
    '                            & "+ Addition" & vbNewLine _
    '                            & "- Subtraction" & vbNewLine _
    '                            & "* Multiply" & vbNewLine _
    '                            & "/ Divide")
    '    SELECT CASE strOperator
    '        CASE "+" , "-" , "*" , "/" 
    '         blIsValid = TRUE
    '    END SELECT
    
    'LOOP

    MsgBox("Selected Operator: " & strOperator)