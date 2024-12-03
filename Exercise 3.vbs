'Scenario: The program should determine if the student can submit assignment 1 
'and assignment 2. This will be based on whether the student scores above 50% 
'for their test. The student needs 55% and above on  assignment 1, 
'if they do, ask them to submit assignment 2. The student needs 60% and above 
'for assignment 2. If the student fails assignment 2, the program should 
'check if they have at least 75% class attendance.

'Declare
    DIM intTestMark, intAttendance, intAssignment1, intAssignment2

'Assign
    
    intTestMark = CInt(InputBox("Please enter test mark"))
    IF intTestMark > 50 THEN '1
        intAssignment1 = CInt(InputBox("Please enter assignment 1 mark"))
        IF intAssignment1 >= 55 THEN '2
             intAssignment2 = CInt(InputBox("Please enter assignment 2 mark"))
            IF intAssignment2 >= 60 THEN '3
                 MsgBox("You passed all the assessments")
            ELSE 
                intAttendance = CInt(InputBox("Please enter class attendance"))
                IF intAttendance >= 75 THEN
                    MsgBox("You have at least 75 attendance")
                ELSE
                    MsgBox("You failed to achieve at least 75% attendance")
                END IF
               
            END IF
        ELSE
            MsgBox("You failed assignment 1, cannot submit assignment 2")

        END IF
    ELSE
        MsgBox("You failed the test, you cannot submit any assignment")

    END IF    
