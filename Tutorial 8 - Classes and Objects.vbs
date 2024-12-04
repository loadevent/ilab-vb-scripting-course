'********************************************************
'          Tutorial 8 - Classes and Objects
'********************************************************
REM This tutorial teaches you about classes and Objects

CLASS Calculator

    public function getWelcomeMessage()
        getWelcomeMessage = "Welcome to our calculator app"
    end function

    public function calcSum(intNum1, intNum2)
        calcSum = (intNum1 + intNum2)
    end function

    public function calcProduct(intNum1, intNum2)
    calcProduct = (intNum1 * intNum2)
    end function

    public function calcDiff(intNum1, intNum2)
        calcDiff = (intNum1 - intNum2)
    end function

    public function calcQuotient(intNum1, intNum2)
        DIM dblQuotient
        IF intNum2 = 0 then
            dblQuotient = "Undefined"
        ELSE
            dblQuotient = intNum1 / intNum2
        END IF
        calcQuotient = dblQuotient
    end function

    public function display()
        display = "Sum: " & sum & vbNewLine _
                    & "Product: " & product & vbNewLine _
                    & "Difference: " & diff & vbNewLine _
                    & "Quotient: " & quotient
    end function

END CLASS
'==============================================================
'Declare
    DIM num1, num2, sum, product, quotient, diff
    Const strTitle = "Calculator App"
    'declare / create an object of Calculator class
    SET objCalculator = NEW Calculator

'Assign
    'call a sub from the class
    MsgBox objCalculator.getWelcomeMessage(), vbOkOnly, strTitle

    num1 = CInt(InputBox("Enter number 1", strTitle))
    num2 = CInt(InputBox("Enter number 2", strTitle))

    sum = objCalculator.calcSum(num1, num2)
    product = objCalculator.calcProduct(num1, num2)
    quotient = objCalculator.calcQuotient(num1, num2)
    diff = objCalculator.calcDiff(num1, num2)
'Use    
    'MsgBox("Sum: " & sum & vbNewLine _
    '                & "Product: " & product & vbNewLine _
    '                & "Difference: " & diff & vbNewLine _
    '                & "Quotient: " & quotient)
    MsgBox objCalculator.display(), vbOkOnly, strTitle

