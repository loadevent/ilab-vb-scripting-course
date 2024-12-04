'display the following menu to the user:

 'Please enter any of these operators:
 '+ Addition
 '- Subtraction
 '* Multiply
 '/ Divide
 
 'the program should allow the user to use one operator from the menu. 
 'Display the menu as long as the user enters an incorrect operator. The program should
 'ask for 2 integers from the user and perform the neccessary computation based on
 'the selected operator, and then display the results.
 
 'The program should as the user if they want to continue using it, if they user says
 'yes then the program should display the menu again, otherwise the program should 
 'exit

 'Declare
    
    DIM strOperator, intNum1, intNum2, dblResults, strResponse
    DO
        'display menu
        DO
            strOperator = InputBox("Please enter any of these operators" & vbNewLine _
                                    & "+ Addition" & vbNewLine _
                                    & "- Subtraction" & vbNewLine _
                                    & "* Multiply" & vbNewLine _
                                    & "/ Divide")
        LOOP UNTIL strOperator = "+" OR strOperator = "-" OR strOperator = "*" OR strOperator = "/"

'Assign
    intNum1 = CInt(InputBox("Please enter number 1"))
    intNum2 = CInt(InputBox("Please enter number 2"))

    SELECT CASE strOperator
        CASE "+" 
            dblResults = intNum1 + intNum2
        CASE "-" 
            dblResults = intNum1 - intNum2
        CASE "*" 
            dblResults = intNum1 * intNum2
        CASE "/" 
           DO WHILE intNum2 = 0
                 intNum2 = CInt(InputBox("Cannot divide by zero. Please enter a different number"))
           LOOP
           dblResults = intNum1 / intNum2
    END SELECT

    MsgBox("Results of " & intNum1 & " " & strOperator & " " & intNum2 & " is: " & dblResults)

    strResponse = LCase(InputBox("Do you want to continue? (Yes/No)"))
    LOOP UNTIL strResponse <> "yes"