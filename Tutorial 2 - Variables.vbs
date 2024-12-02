'********************************************************
'                     Tutorial 2 - Variables
'********************************************************
REM This tutorial teaches you about different data types
Option Explicit

'Declare
DIM strName
DIM intAge
DIM dblHeight
DIM isMarried

'Assign
'Get input from the user
strName = InputBox("Please enter your name")
intAge = CInt(InputBox("Please enter you age"))
dblHeight = CDbl(InputBox("Please enter your height"))
isMarried = CBool(CInt(InputBox("Is Married? " & vbNewLine _
                                            & "1 - Yes" & vbNewLine _
                                            & "0 - No")))


'Use / Consume
MsgBox("Name: " & strName & vbNewLine _
                & "Age: " & intAge & vbNewLine _
                & "Height: " & dblHeight & vbNewLine _
                & "Is Married? " & isMarried)