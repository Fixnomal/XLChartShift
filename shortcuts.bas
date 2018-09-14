Attribute VB_Name = "shortcuts"
Function qr(firstCellRow As Long, firstCellColumn As Long, Optional lastCellRow As Long, Optional lastCellColumn As Long, _
 Optional ws As Worksheet, Optional explicitErrorMessage As Boolean = False) As Range
'returns range from 2 or 4 longs, optional inclusion of worksheet

    Dim returnToWs As Worksheet
    
    If ws Is Nothing Then
        If lastCellRow = 0 And lastCellColumn = 0 And firstCellRow > 0 And firstCellColumn > 0 Then
            Set qr = Range(Cells(firstCellRow, firstCellColumn), Cells(firstCellRow, firstCellColumn))
        ElseIf lastCellRow > 0 And lastCellColumn > 0 And firstCellRow > 0 And firstCellColumn > 0 Then
            Set qr = Range(Cells(firstCellRow, firstCellColumn), Cells(lastCellRow, lastCellColumn))
        Else
            If explicitErrorMessage = True Then
                MsgBox ("Error in qr function: Values for cells and rows must be greater than 0")
            End If
        End If
    Else
        Set returnToWs = ActiveSheet
        ws.Activate
        If lastCellRow = 0 And lastCellColumn = 0 And firstCellRow > 0 And firstCellColumn > 0 Then
            Set qr = ws.Range(Cells(firstCellRow, firstCellColumn), Cells(firstCellRow, firstCellColumn))
        ElseIf lastCellRow > 0 And lastCellColumn > 0 And firstCellRow > 0 And firstCellColumn > 0 Then
            Set qr = ws.Range(Cells(firstCellRow, firstCellColumn), Cells(lastCellRow, lastCellColumn))
        Else
            If explicitErrorMessage = True Then
                MsgBox ("Error in qr function: Values for cells and rows must be greater than 0")
            End If
        End If
        returnToWs.Activate
    End If
End Function


Sub subStart()
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Sub subEnd()
    'Last modified: 20140408
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
End Sub
