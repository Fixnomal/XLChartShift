VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChartCopyForm 
   Caption         =   "Smart Chart Copy"
   ClientHeight    =   4440
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   5364
   OleObjectBlob   =   "ChartCopyForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChartCopyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Done_Click()
    'Last modified: 20140329
    
    Dim templateChart As Chart
    
    'Check if Chart selected
    Set templateChart = ActiveChart
    If templateChart Is Nothing Then
        MsgBox ("Select a chart to be copied before starting the Macro")
        Unload Me
        Exit Sub
    End If
    
    'Call Chart Copy sub with user input values
    Call chartCopy(templateChart, CopNum.Value, LocRowOff.Value, LocColOff.Value, TitRowOff.Value, _
    TitColOff.Value, titOff.Value, YRowOff.Value, YColOff.Value, XRowOff.Value, _
    XColOff.Value)
    
    Unload Me
    
End Sub

Private Sub Cncl_Click()
    Unload Me
End Sub


Private Sub Frame4_Click()

End Sub

Private Sub UserForm_Initialize()

    'Set Defaults
    LocRowOff.Value = 24
    LocColOff.Value = 0
    TitRowOff.Value = 0
    TitColOff.Value = 0
    YRowOff.Value = 24
    YColOff.Value = 0
    ErrBar.Value = False
    titOff.Value = True
    CopNum.Value = 1
    XRowOff.Value = 0
    XColOff.Value = 0
    
End Sub

Private Sub LocRowOff_Change()
    YRowOff.Value = LocRowOff.Value
End Sub

Private Sub LocColOff_Change()
    YColOff.Value = LocColOff.Value
End Sub


