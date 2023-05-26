Attribute VB_Name = "Module1"
Sub Resize_Table()
Attribute Resize_Table.VB_ProcData.VB_Invoke_Func = " \n14"
'
' This function resizes and reorders the table based on the number of study sessions.
'

    Dim Location As String
    Dim userInput As Integer
    Dim x As Integer
    Dim timeInput As String
    Dim locationValue As String
    Dim Need As String
    Dim tryAgain As Integer
    
    
    userInput = InputBox("How many sessions do you want?")
    x = 1
    
    
    Go_to_Bottom
    
   'Loop until x reaches the total number of iterations based on user input
    Do While x < (userInput * 4) + userInput
    
        'Move the active cell one row below
        ActiveCell.Offset(1, 0).Select
        
        'Increment the counter x by 1
        x = x + 1
        
        
        
    Loop
    
    
    
    'Store the address of the active cell in the "Location" variable
    Location = ActiveCell.Address
    
    
    'Resize the range of the "AutoPomodoro" list object to include cells from A1 to the address stored in the "Location" variable
    ActiveSheet.ListObjects("AutoPomodoro").Resize Range("$A$1:" + Location)
    
    
    Selection.End(xlToLeft).Select
    
    ' In the case where the user desires multiple sessions, we follow a different path.
    If userInput > 1 Then
        x = 1
        
        'Check if userInput equals 2
        If userInput = 2 Then
        
            'Execute the loop while x is less than 5
            Do While x < 5
                
                ActiveCell.Offset(-1, 0).Select
                x = x + 1
            Loop
            
        'Set the value of the active cell as "12:15:00 AM"
        ActiveCell.FormulaR1C1 = "12:15:00 AM"
        
        'Move the active cell one column to the right
        ActiveCell.Offset(0, 1).Select
        
        'Set the value of the active cell as "BREAK"
        ActiveCell.Value = "BREAK"
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Offset(0, -1).Select
        
        'Set the formula of the active cell as "=RC[-1]+M"
        'M(R18) is the break time betwwen sessions which is 15 minutes.
        ActiveCell.FormulaR1C1 = "=RC[-1]+M"
        ActiveCell.Offset(0, 1).Select
        
        End If
        
        x = 1
        
        'Check if userInput is greater than 2
        If userInput > 2 Then
        
            'Execute the loop while x is less than the calculated value based on userInput
            Do While x < userInput * 3 + (userInput - 3) * 2 + 2
            
                ActiveCell.Offset(-1, 0).Select
                x = x + 1
                
                'Check if x is a multiple of 5 and greater than 5
                If (x - 5) Mod 5 = 0 And x > 5 Then
                
                    'Set the value of the active cell as "12:15:00 AM"
                    ActiveCell.FormulaR1C1 = "12:15:00 AM"
                    
                    
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = "BREAK"
                    ActiveCell.Offset(0, -1).Select
                    ActiveCell.Offset(0, -1).Select
                    ActiveCell.FormulaR1C1 = "=RC[-1]+M"
                    ActiveCell.Offset(0, 1).Select
                
                'Check if x equals 5
                ElseIf x = 5 Then
                    ActiveCell.FormulaR1C1 = "12:15:00 AM"
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = "BREAK"
                    ActiveCell.Offset(0, -1).Select
                    ActiveCell.Offset(0, -1).Select
                    
                    'Set the formula of the active cell as "=RC[-1]+M"
                    ActiveCell.FormulaR1C1 = "=RC[-1]+M"
                    ActiveCell.Offset(0, 1).Select
                    
                End If
                
        
                
            Loop
        End If
    End If
    
    'To apply the Pomodoro time schedule, we navigate to the cell containing the start time and copy the address of that cell.
    'Move the active cell one column to the left
    ActiveCell.Offset(0, -1).Select
    'Move the active cell one more column to the left
    ActiveCell.Offset(0, -1).Select
    'Select the last cell in the column by moving up from the active cell until a non-empty cell is encountered
    Selection.End(xlUp).Select
    'Move the active cell one row below
    ActiveCell.Offset(1, 0).Select
    'Store the address of the active cell in the "Location" variable
    Location = ActiveCell.Address
    
    
    
    Copy_Session_Pattern
    'Copy the session pattern by selecting the range specified by the "Location" variable
    Range(Location).Select
    'Paste the copied session pattern in the active cell using the PasteSpecial method
    ActiveCell.PasteSpecial
    'Select the active cell after pasting the session pattern
    ActiveCell.Select
    
    'Prompt the user to input the desired start time using an InputBox dialog
    timeInput = InputBox("What time do you wanna start?")
    'Set the value of the active cell to the user-inputted start time.The remaining cells used for displaying times will be updated automatically.
    ActiveCell.Value = timeInput
    
    
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Copy 'To set the starting time of the break, we copy the value of the active cell.
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(1, 0).Select
    
    If userInput = 1 Then 'We don't need to set break time since it is only one session.
        
        'Disable CutCopyMode to clear any cut or copied data
        Application.CutCopyMode = False
        'Enable screen updating to refresh the display
        Application.ScreenUpdating = True
        'Exit the subroutine and continue with the rest of the code
        Exit Sub
        
    End If
    
    ' Perform a paste special operation to paste values and number formats without any additional operations
    Selection.PasteSpecial PASTE:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, TRANSPOSE:=False
    
    ActiveCell.Offset(0, 1).Select
    
    'Store the address of the active cell in the variable 'Need'
    Need = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(1, 0).Select
    
    'Store the address of the current active cell in the variable 'Location'
    Location = ActiveCell.Address
    
    i = 1
    
    'Call the 'Copy_Session_Pattern' function
    Copy_Session_Pattern
    
    'Select the range specified by the 'Location' address
    Range(Location).Select
    
    'Perform a paste special operation to paste the copied content in the active cell
    ActiveCell.PasteSpecial
    
    
    ActiveCell.Select
    
    'Select the range specified by the 'Need' address
    Range(Need).Select
    Range(Need).Copy
    Range(Location).Select
    
    'Perform a paste special operation to paste the copied content in the active cell
    Selection.PasteSpecial PASTE:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, TRANSPOSE:=False
    
    ActiveCell.Select
    
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(0, 1).Select
    
    'Check if userInput equals 2
    If userInput = 2 Then
    
        'Clear the clipboard
        Application.CutCopyMode = False
        
        'Select the cell one column to the left of the current active cell
        ActiveCell.Offset(0, -1).Select
        
        'Select the cell one row below the current active cell
        ActiveCell.Offset(1, 0).Select
        
        'Enable screen updating
        Application.ScreenUpdating = True
        
        Exit Sub
        
    End If
    
    ActiveCell.Copy
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(1, 0).Select
    
    'Perform a paste special operation to paste values and number formats without any additional operations
    Selection.PasteSpecial PASTE:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, TRANSPOSE:=False
    ActiveCell.Offset(0, 1).Select
    
    'Store the address of the active cell in the variable 'Need'
    Need = ActiveCell.Address
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(1, 0).Select
    
    'Store the address of the current active cell in the variable 'Location'
    Location = ActiveCell.Address
    
    'Execute the loop while i is less than userInput - 1
    Do While i < userInput - 1
        'Call the 'Copy_Session_Pattern' function
        Copy_Session_Pattern
        
        'Select the range specified by the 'Location' address
        Range(Location).Select
        
        'Perform a paste special operation to paste the copied content in the active cell
        ActiveCell.PasteSpecial
        
        'Select the range specified by the 'Need' address
        Range(Need).Select
        'Copy the range specified by 'Need'
        Range(Need).Copy
        
        'Select the range specified by the 'Location' address
        Range(Location).Select
        
        'Perform a paste special operation to paste the copied content in the active cell
        Selection.PasteSpecial PASTE:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, TRANSPOSE:=False
        
        ActiveCell.Select
        
        
        
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Copy
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Offset(1, 0).Select
        
        'Check if i is equal to userInput - 2
        If i = userInput - 2 Then
        
            'Exit the loop
            Exit Do
        End If
        
        'Perform a paste special operation to paste values and number formats without any additional operations
        Selection.PasteSpecial PASTE:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, TRANSPOSE:=False
        
        
        ActiveCell.Offset(0, 1).Select
        
        'Store the address of the active cell in the variable 'Need'
        Need = ActiveCell.Address
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Offset(1, 0).Select
        
        'Store the address of the current active cell in the variable 'Location'
        Location = ActiveCell.Address
        
        'Increment the value of 'i' by 1
        i = i + 1
    Loop
    
    'Clear the clipboard
    Application.CutCopyMode = False
                
    
End Sub
Sub Go_to_Bottom()
'
' Bottom_of_Table Macro
' Find the bottom of table in the result column.

' In this step, we will select the first row in the "Result" column and navigate to the bottom of the column.

    Range("F2").Select
    Selection.End(xlDown).Select
    
    
End Sub
Sub Copy_Session_Pattern()
'
' Copy_Session_Pattern Macro
' It copies the session pattern
'

'
    Range("O15:P18").Select
    Selection.Copy
End Sub
Sub Writing_Dates()
Attribute Writing_Dates.VB_ProcData.VB_Invoke_Func = " \n14"
'
' It writes today's date to the Date Cells
'

'
    
    Dim Loc As String
    Loc = ActiveCell.Address
    
    'Copy the value from cell P3
    Range("P3").Select
    Selection.Copy
    
    'Go to the last used cell in the same column as the active cell
    Range(Loc).Select
    Selection.End(xlUp).Select
    
    'Go to the last used cell in the same row as the previously selected cell
    Selection.End(xlToLeft).Select
    
    'Go up one cell to the row above
    Selection.End(xlUp).Select
    
    'Go down one cell to the next empty row
    ActiveCell.Offset(1, 0).Select
    
    'Select the range from the current cell to the last used cell in the column
    Range(Selection, Selection.End(xlDown)).Select
    
    'Paste the copied value, including number formats, to the selected range
    Selection.PasteSpecial PASTE:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, TRANSPOSE:=False
        
    'Clear the clipboard
    Application.CutCopyMode = False
    
    'Select the active cell and go to the next cell in the same row
    ActiveCell.Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    
End Sub

Sub Checking()
'
' Find the bottom of table in "RESULT" column to check if it's empty
'

'
    
    
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    
    
    
End Sub

Public Sub Without_Pomodoro()
Attribute Without_Pomodoro.VB_Description = "If you prefer to study without implementing the Pomodoro technique and without a specific plan, you have the option to utilize this keyboard shortcut."
Attribute Without_Pomodoro.VB_ProcData.VB_Invoke_Func = "p\n14"


' This subprocedure creates a single-row study plan without using the Pomodoro technique and without an ending time.
' It is intended for those who wish to study without a specific plan.

    Dim Loca As String
    Dim Inp As String
    
    'Call the REMINDER subroutine or function from Module3
    Module3.REMINDER
    
    ' Call the Checking subroutine or function from Module1
    Module1.Checking
    
    
    ' Check if the active cell is empty
    If ActiveCell.Value = "" Then
    
        'Display a message box to inform the user to complete the empty cells
        MsgBox ("First complete the empty cells!")
        Exit Sub
    End If
    
    ' Call the Go_to_Bottom function
    Go_to_Bottom
    
    ActiveCell.Offset(1, 0).Select
    
    ' Store the address of the current active cell in the variable 'Loca'
    Loca = ActiveCell.Address
    
    'Resize the "AutoPomodoro" list object to include the range from cell A1 to the stored 'Loca' address
    ActiveSheet.ListObjects("AutoPomodoro").Resize Range("$A$1:" + Loca)
    
    ' Set the value of the "RESULT" cell to "+"
    ActiveCell.Value = "+"
    
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(0, -1).Select
    
    'Clear the value of the cell three columns to the left of the current active cell
    ActiveCell.Value = ""
    ActiveCell.Offset(0, -1).Select
    
    ' Prompt the user to enter a start time and store the input in the variable 'Inp'
    Inp = InputBox("What time do you wanna start?")
    ' Set the value of the current active cell to the stored input 'Inp'
    ActiveCell.Value = Inp
    
    ActiveCell.Offset(0, -1).Select
    
    ' Store the address of the current active cell in the variable 'Loca'
    Loca = ActiveCell.Address
    
    ' Select cell P3(Today's date)
    Range("P3").Select
    Selection.Copy
    
    ' Select the range specified by the 'Loca' address
    Range(Loca).Select
    
    ' Perform a paste special operation to paste the copied content in the active cell
    Selection.PasteSpecial PASTE:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, TRANSPOSE:=False
        
    'Clear the clipboard
    Application.CutCopyMode = False
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    
    'Call the RefreshAndMovePivotTable function from Module3
    Module3.RefreshAndMovePivotTable
    
    ActiveCell.Offset(-1, 0).Select
    
    
    
End Sub

