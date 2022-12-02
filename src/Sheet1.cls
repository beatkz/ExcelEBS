VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tgtRow As Long
    Const COL_TSK_NUM As Long = 1
    Const COL_END_TIME As Long = 4
    Dim t As Range
    Dim TSheetWs As Worksheet
    Set TSheetWs = ThisWorkbook.Worksheets("TSheet")
    If Target.column = COL_TSK_NUM Then
        For Each t In Target
            tgtRow = t.Row
            'Write Task Name
            If t.Value <> "" Then
                '=IFERROR(VLOOKUP(A3,Tasks!B:D,3,FALSE),VLOOKUP(A3,OldTasks!B:D,3,FALSE))
                TSheetWs.Range("G" & tgtRow).Formula = _
                "=IFERROR(VLOOKUP(A" & tgtRow & ",Tasks!B:D,3,FALSE)," & _
                "VLOOKUP(A" & tgtRow & ",OldTasks!B:D,3,FALSE))"
            End If
        Next
    ElseIf Target.column = COL_END_TIME Then
        For Each t In Target
            tgtRow = t.Row
            'Write Time Spent in Hours
            TSheetWs.Range("E" & tgtRow).Formula = "=(D" & tgtRow & "-C" & tgtRow & ")*24"
            
            'Write Time Spent in Minutes
            TSheetWs.Range("F" & tgtRow).Formula = "=E" & tgtRow & "*60"
        Next
    End If
    Set TSheetWs = Nothing
End Sub
