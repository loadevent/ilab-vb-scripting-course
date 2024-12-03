'Create a program that will receive the fullname from the user, 
'split the fullname into firstname and lastname. Then display 
'firstname and lastname separately, the first letters of the 
'firstname and lastname should be in upper case

'input
'fullname: mike ross

'output
'firstname: Mike
'lastname: Ross

'Declare
DIM strFullname
DIM strFirstname
DIM strLastname
DIM intSpacePos
'mike ross
'michael jordan
'jessica jones

'Assign
strFullname = Trim(InputBox("Please enter fullname"))
intSpacePos = InStr(strFullname, " ")
'm  i  k  e   r  o  s  s
'1  2  3  4 5 6  7  8  9
strFirstname = Trim(Left(strFullname, intSpacePos -1))
strLastname = Trim(Right(strFullname, Len(strFullname) - intSpacePos))

strFirstname = UCase(Left(strFirstname, 1)) & LCase(Right(strFirstname, Len(strFirstname) -1))
strLastname = UCase(Left(strLastname, 1)) & LCase(Right(strLastname, Len(strLastname) -1))

'Use
MsgBox("Firstname: " & strFirstname & vbNewLine _
                    & "Lastname: " & strLastname)