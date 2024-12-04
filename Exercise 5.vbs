'create a program that will perform different calculations using 
'functions. The program should have the following functions:
'1 - a function to calculate the sum 
'2 - a function to calculate the product 
'3 - a function to calculate the difference 
'4 - a function to calculate the quotient
'All these functions should take 2 numbers as parameters 
'The program should have the following subroutine:
'1 - a sub to display a welcome message
'Lastly, the program should display the sum, difference, product
'and quotient of the 2 numbers

sub welcomeMessage()
    MsgBox("Welcome to our basic calculation universe")
end sub

function calcSum(intNum1, intNum2)
    DIM intSum
    intSum = CInt(intNum1) + CInt(intNum2)
    calcSum = intSum
end function

function calcProduct(intNum1, intNum2)
    calcProduct = (intNum1 * intNum2)
end function

function calcDiff(intNum1, intNum2)
    calcDiff = (intNum1 - intNum2)
end function

function calcQuotient(intNum1, intNum2)
    DIM dblQuotient
    IF intNum2 = 0 then
        dblQuotient = "Undefined"
    ELSE
        dblQuotient = intNum1 / intNum2
    END IF
    calcQuotient = dblQuotient
end function

'===============================================================================
    DIM intNum1, intNum2
    DIM sum, product, difference, quotient
    DIM strResponse

    call welcomeMessage()

    DO

    intNum1 = CInt(InputBox("Please enter the first number"))
    intNum2 = CInt(InputBox("Please enter the second number"))

    'DO WHILE intNum2 = 0
    '     intNum2 = CInt(InputBox("Cannot divide by zer. Please enter a different number to divide with"))
    'LOOP

    sum = calcSum(intNum1, intNum2)
    product = calcProduct(intNum1, intNum2)
    difference = calcDiff(intNum1, intNum2)

    quotient = calcQuotient(intNum1, intNum2)

    MsgBox("The sum of " & intNum1 & " and " & intNum2 & " is: " & sum & vbNewLine _
                & "The product of " & intNum1 & " and " & intNum2 & " is: " & product & vbNewLine _
                & "The difference of " & intNum1 & " and " & intNum2 & " is: " & difference & vbNewLine _
                & "The quotient of " & intNum1 & " and " & intNum2 & " is: " & quotient)

    strResponse = Trim(LCase(InputBox("Do you want to exit? (Y/N)")))

    LOOP UNTIL strResponse = "y"