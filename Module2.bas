Attribute VB_Name = "Module2"
Public Sub Main()
Attribute Main.VB_Description = "To utilize the Pomodoro technique for studying, you can employ this keyboard shortcut."
Attribute Main.VB_ProcData.VB_Invoke_Func = "j\n14"

' This is the primary function that is triggered when someone wants to study using the Pomodoro technique.




    'By turning off screen updating, Excel avoids constantly refreshing the screen during the execution of the macro,
    'which can significantly speed up the code execution and prevent unnecessary flickering or visual distractions for the user.
    Application.ScreenUpdating = False
    
    'The code enters the Module3 REMINDER function
    Module3.REMINDER
    
    'Call the "Checking" subroutine in Module1
    Module1.Checking
    
    'Check if the active cell is empty
    If ActiveCell.Value = "" Then
        MsgBox ("First complete the empty cells!")
        Exit Sub
    End If
    
        
    'Call the "Resize_Table" subroutine in Module1
    Module1.Resize_Table
    
    'Call the "Writing_Dates" subroutine in Module1
    Module1.Writing_Dates
    
    'Call the "RefreshAndMovePivotTable" subroutine in Module3
    Module3.RefreshAndMovePivotTable
    
    'Enable screen updating
    Application.ScreenUpdating = True
End Sub
