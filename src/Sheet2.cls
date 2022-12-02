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
    Const COL_TASK As Long = 2
    Dim t As Range
    
    'Workaround: When you select the entire row, skip this.
    If Target.Count >= 4096 Then
        Exit Sub
    End If
    
    Select Case Target.column
    Case COL_STAT
        For Each t In Target
            tgtRow = t.Row
            If t.Value = "Ï" Or t.Value = "Done" Then
                Range("D" & tgtRow, "G" & tgtRow).Font.Strikethrough = True
            Else
                Range("D" & tgtRow, "G" & tgtRow).Font.Strikethrough = False
            End If
        Next
    Case COL_TASK
        Dim TaskWs As Worksheet
        Set TaskWs = ThisWorkbook.Worksheets("Tasks")
        For Each t In Target
            tgtRow = t.Row
            
            Range("C" & tgtRow).Validation.Delete
            With Range("C" & tgtRow).Validation
                .Add Type:=xlValidateList, _
                    Formula1:="=Projects!$A$3:$A$15"
            End With
            
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
            "=IF(H" & tgtRow & ">0,ROUNDUP(G" & tgtRow & "/H" & tgtRow & ",2),0)"
        Next
        Set TaskWs = Nothing
    Case Else
    End Select
End Sub
