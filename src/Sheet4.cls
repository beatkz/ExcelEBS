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
    'For Est vs Act Chart
    maxRow = EBS.getMaxRow("Tasks", Range("B2").column)
    ActiveSheet.ChartObjects(1).Select
    ActiveChart.SetSourceData Source:=Sheets( _
        "Tasks").Range( _
        "G2:G" & CStr(maxRow) & ",H2:H" & CStr(maxRow))
    'For Confidence Distro Chart
    ActiveSheet.ChartObjects(2).Select
    ActiveChart.SetSourceData Source:=Sheets( _
        "Sim").Range( _
        "F3:DB4")
    'For Ship Date Distro Chart
    maxRow = EBS.getMaxRow("ShipDateLog", Range("A2").column)
    ActiveSheet.ChartObjects(3).Select
    ActiveChart.SetSourceData Source:=Sheets( _
        "ShipDateLog").Range( _
        "A2:D" & CStr(maxRow))
    
    
    Application.ScreenUpdating = True
End Sub
