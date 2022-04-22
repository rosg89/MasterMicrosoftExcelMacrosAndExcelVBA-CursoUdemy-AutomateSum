Attribute VB_Name = "Module1"
Sub AutomateSum()
    
    Dim lastCell As String
    Dim i As Integer
    
    i = 1
    
    'loop
    Do While i <= Worksheets.Count
    Worksheets(i).Select
    
    'selects the F2 cell
    Range("F2").Select
    
    'selects the last cell in the column
    Selection.End(xlDown).Select
    
    lastCell = ActiveCell.Address(False, False)
    
    'selects la de abajo
    ActiveCell.Offset(1, 0).Select
    
    'sum
    ActiveCell.Value = "=SUM(F2:" & lastCell & ")"
    
    i = i + 1
    
    Loop
    
End Sub
