Attribute VB_Name = "Module3"
Sub RefreshAndMovePivotTable()
Attribute RefreshAndMovePivotTable.VB_ProcData.VB_Invoke_Func = " \n14"
'
' It updates the information in the pivot table and brings it to a visible location.
'

'
    'Refresh the pivot cache and update the PivotTable2
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    
    'Select the entire PivotTable2, including data and labels
    ActiveSheet.PivotTables("PivotTable2").PivotSelect "", xlDataAndLabel, True
    
    'Cut the selected PivotTable2
    Selection.Cut
    
    'Go to the bottom of the table
    Module1.Go_to_Bottom
    
    'Go to the cell one column to the right
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Offset(0, 1).Select
    
    'Paste the cut PivotTable2
    ActiveSheet.PASTE
    
    'Select the active cell
    ActiveCell.Select
    
    'Go to the cell one column to the left
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Offset(0, -1).Select
    
    'Go to the cell one row down
    ActiveCell.Offset(1, 0).Select
    
    
    
End Sub
Public Sub REMINDER()

'This is a public subroutine named REMINDER
'It displays a message box reminding the user to read a book


    MsgBox ("DON'T FORGET READING BOOK!")
    
    
End Sub
