Attribute VB_Name = "EBS"
Option Explicit
Sub SimulateFuture()
    Dim SheetName As String
    SheetName = ActiveSheet.Name
    Dim maxRow As Long
    maxRow = getMaxRow(SheetName, Range("A1").column)
    
    Dim currentRow As Long
    currentRow = 8
    
    Dim col As Long
    
    'Pause Auto Calculation
    Application.Calculation = xlCalculationManual
    
    Do While currentRow <= maxRow
        For col = 6 To 105
            Cells(currentRow, col) = _
            "=" & Cells(currentRow, "E") & _
            "/" & randomPickedVelocityfromDoneTasks()
        Next col
        currentRow = currentRow + 1
    Loop
    
    'Resume Auto Calculation
    Application.Calculation = xlCalculationAutomatic
End Sub

Function randomPickedVelocityfromDoneTasks()
    Const MIN_ROW As Long = 3
    Dim maxRow As Long
    maxRow = getMaxRow("Tasks", Range("A1").column)
    
    Dim velocity As Double
    Dim pickedRow As Long
    Randomize
    pickedRow = Int((maxRow - MIN_ROW) * Rnd) + MIN_ROW
    velocity = Sheets("Tasks").Cells(pickedRow, Range("I3").column)
    
    randomPickedVelocityfromDoneTasks = velocity
End Function
Function getMaxRow(SheetName As String, column As Long) As Long
    getMaxRow = Sheets(SheetName).Cells(Rows.Count, column).End(xlUp).Row
End Function


