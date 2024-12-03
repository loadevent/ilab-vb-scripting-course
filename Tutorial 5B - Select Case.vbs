
'Declare
    DIM intDayNumber
    DIM strDay

'Assign
    intDayNumber = 13

    SELECT CASE intDayNumber
        CASE 1 
            strDay = "Monday"
        CASE 2 
            strDay = "Tuesday"
        CASE 3
            strDay = "Wednesday"
        CASE 4 
            strDay = "Thursday"
        CASE 5 strDay = "Friday"
        CASE 6 strDay = "Saturday"
        CASE 7 strDay = "Sunday"
        CASE ELSE strDay = "Incorrect number: [" & intDayNumber & "]" 
    END SELECT

    MsgBox(strDay)