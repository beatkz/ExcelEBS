VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Dim maxRow As Long
    maxRow = EBS.getMaxRow("Tasks", Range("B2").column)
    ActiveSheet.ChartObjects(1).Select
    ActiveChart.SetSourceData Source:=Sheets( _
        "Tasks").Range( _
        "F2:F" & CStr(maxRow) & ",H2:H" & CStr(maxRow))
    Application.ScreenUpdating = True
End Sub
