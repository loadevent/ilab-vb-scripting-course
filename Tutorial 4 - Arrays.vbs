'********************************************************
'             Tutorial 4 - Arrays
'********************************************************
REM This tutorial teaches you about Arrays
REM Array - variable that can store multiple values of the same data types

'Declare
    DIM arIntNumbers(5) 'empty array to store 5 items / numbers
    'delcare an array with initial values
    DIM arStrNames

    arStrNames = Array("Kelvin", "Thabo", "Kabelo", "Caroline", "David", "Jessica")

'Assign
    arIntNumbers(0) = 45
    arIntNumbers(1) = 38
    arIntNumbers(2) = 75
    arIntNumbers(3) = 25
    arIntNumbers(4) = 55

'Use
'Access array elements
    MsgBox(arIntNumbers(2))
    MsgBox(arStrNames(5))

