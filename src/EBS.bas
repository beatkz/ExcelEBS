Attribute VB_Name = "EBS"
Option Explicit
Sub SimulateFuture()
    Dim SheetName As String
    SheetName = ActiveSheet.Name
    
    Dim currentRow As Long
    currentRow = 8
    
    Dim col As Long
    
    Range("A8").Select
    
    Call init
    Call fillUndoneTaskNo
    
    Dim maxRow As Long
    maxRow = getMaxRow(SheetName, Range("A1").column)
    
    'Pause Auto Calculation
    Application.Calculation = xlCalculationManual
    
    Do While currentRow <= maxRow
        'Fill VLOOKUP formulas
        With Range("B" & currentRow & ":C" & currentRow)
            .WrapText = True
        End With
        Range("B" & currentRow).Formula = _
        "=IFERROR(VLOOKUP(A" & currentRow & ",Tasks!$B$2:$I$102,2,FALSE),"""")"
        Range("C" & currentRow).Formula = _
        "=IFERROR(VLOOKUP(A" & currentRow & ",Tasks!$B$2:$I$102,3,FALSE),"""")"
        Range("D" & currentRow).Formula = _
        "=IFERROR(VLOOKUP(A" & currentRow & ",Tasks!$B$2:$I$102,5,FALSE),"""")"
        Range("E" & currentRow).Formula = _
        "=IFERROR(VLOOKUP(A" & currentRow & ",Tasks!$B$2:$I$102,6,FALSE),"""")"

        For col = 6 To 105
            Cells(currentRow, col) = _
            "=" & Cells(currentRow, "E") & _
            "/" & randomPickedVelocityfromTasks()
        Next col
        currentRow = currentRow + 1
    Loop
    
    'Resume Auto Calculation
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub init()
    '--- Cell font Config ---
    Const FONT_NAME As String = "Meiryo UI"
    Const FONT_SIZE As Integer = 8
    '--- End Cell font Config ---
    
    'Clear all simulated times
    Range("F8:DA108").Clear
    'Clear all undone task numbers
    Range("A8:E108").Clear
    
    With Range("A1:DA108").Font
        .Name = FONT_NAME
        .Size = FONT_SIZE
    End With
    
    ' Total Hours formula:
    ' "=SUM(F8:F108)"
    ' Total Days formula:
    ' "=ROUNDUP(F2/$B$4,2)"
    ' Probability formula:
    ' "=NORM.DIST(F3,AVERAGE($F$3:$DA$3),STDEV.S($F$3:$DA$3),TRUE)"
End Sub

Sub fillUndoneTaskNo()
    Application.ScreenUpdating = False
    
    Sheets("Tasks").Select

    ActiveSheet.Range("$A$2:$F$4096").AutoFilter Field:=1, Criteria1:="="
    ActiveWindow.SmallScroll Down:=-9
    Range("B3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sim").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A8").Select
    Application.CutCopyMode = False
    Sheets("Tasks").Select
    ActiveSheet.ShowAllData
    Cells(getMaxRow("Tasks", 1) + 1, 1).Select
    
    Sheets("Sim").Select
    Application.ScreenUpdating = True
End Sub

Function randomPickedVelocityfromTasks() As Double
    Dim maxRow As Long
    maxRow = getMaxRow("Tasks", Range("A1").column)
    
    Dim velocity As Double
    velocity = drawVelocity(maxRow)
    
    randomPickedVelocityfromTasks = velocity
End Function

Function drawVelocity(maxRow) As Double
    Const MIN_ROW As Long = 3
    Dim pickedRow As Long
    Dim velocity As Double

Redraw:
    Randomize
    pickedRow = Int((maxRow - MIN_ROW) * Rnd) + MIN_ROW
    velocity = CDbl(Sheets("Tasks").Cells(pickedRow, Range("I3").column))
    If velocity = 0# Then GoTo Redraw
    
    drawVelocity = velocity
End Function

Function getMaxRow(SheetName As String, column As Long) As Long
    getMaxRow = Sheets(SheetName).Cells(Rows.Count, column).End(xlUp).Row
End Function
