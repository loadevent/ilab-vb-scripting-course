'Create a program that will receive firstname and lastname
'from the user. Then display the fullname such that the first
'letter of the firstname and lastname are capital letters

'input: 
'firstname: david
'lastname: smith
'-----------------------
'firstname: CAROLINE
'lastname: jones

'output:
'David Smith
'-----------------------
'Caroline Jones

'Declare
DIM strFirstname, strLastname
DIM strFullname

'Assign
strFirstname = Trim(InputBox("Please enter firstname"))
strLastname = Trim(InputBox("Please enter lastname"))

strFirstname = LCase(strFirstname)
strLastname = LCase(strLastname)

strFullname = UCase(Left(strFirstname,1)) & Right(strFirstname,Len(strFirstname) -1) & " " + UCase(Left(strLastname,1)) + Right(strLastname,Len(strLastname) -1)

'Use
MsgBox("Fullname: " & strFullname)