CLASS Employee
    'declare class variables
    public strFirstname
    public strLastname
    public strEmployeeNo
    public intHoursWorked
    
    
    public function getWelcomeMessage()
        getWelcomeMessage = "Hi " & strFirstname & " " & strLastname & ", Welcome to the Employee Management System"
    end function

    private function calcSalary()
        Const dblRate = 255
        calcSalary = (intHoursWorked * dblRate)
    end function

    private function getEmployeeNo()
        getEmployeeNo = "EMP-" & strEmployeeNo
    end function

    public function getDetails()
        getDetails = "Fullname: " & strFirstname & " " & strLastname & vbNewLine _
                                    & "Employee No: " & getEmployeeNo() & vbNewLine _
                                    & "Hours Worked: " & intHoursWorked & vbNewLine _
                                    & "Salary: " & calcSalary()
    end function
End class
'=====================================================================
'Declare
    SET objEmployee =  NEW Employee
    Const strTitle = "Employee Management System"

'Assign
    objEmployee.strEmployeeNo = InputBox("Please enter your employee no", strTitle)
    objEmployee.strFirstname = InputBox("Please enter firstname", strTitle)
    objEmployee.strLastname = InputBox("Please enter lastname", strTitle)
    objEmployee.intHoursWorked = CInt(InputBox("Please enter hours worked", strTitle))
    

'Use
    MsgBox objEmployee.getDetails(), vbOkOnly, strTitle

