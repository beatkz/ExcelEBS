VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tgtRow As Long
    Const COL_STAT As Long = 1
    Const COL_PROJ As Long = 3
    Dim t As Range
    If Target.column = COL_STAT Then
        For Each t In Target
            tgtRow = t.Row
            If t.Value = "Ï" Or t.Value = "Done" Then
                Range("D" & tgtRow, "G" & tgtRow).Font.Strikethrough = True
            Else
                Range("D" & tgtRow, "G" & tgtRow).Font.Strikethrough = False
            End If
        Next
    ElseIf Target.column = COL_PROJ Then
        Dim TaskWs As Worksheet
        Set TaskWs = ThisWorkbook.Worksheets("Tasks")
        For Each t In Target
            tgtRow = t.Row
            'Write Proj.Name
            Dim CellVal As String
            CellVal = "=VLOOKUP(C" & tgtRow & ",Projects!A:C,3,FALSE)"
            TaskWs.Range("G" & tgtRow).Formula = CellVal
            
            'Write Actual Hour(s)
            Dim sumActHour As String
            sumActHour = "SUMIF(" _
                & "TSheet!A:E," _
                & "Tasks!B" & tgtRow & "," _
                & "TSheet!E:E)"
                
            TaskWs.Range("H" & tgtRow).Formula = _
            "=IF(" & sumActHour & ">0," _
            & "ROUNDUP(" & sumActHour & ",2),0)"
            
            'Write Velocity
            TaskWs.Range("I" & tgtRow).Formula = _
            "=IF(H" & tgtRow & ">0,ROUNDUP(F" & tgtRow & "/H" & tgtRow & ",2),0)"
        Next
        Set TaskWs = Nothing
    End If
End Sub
