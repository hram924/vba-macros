VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuarterForm 
   Caption         =   "Quarter Form"
   ClientHeight    =   2863
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "vbaform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuarterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VBA Form'

Private Sub cboWhichSheet_Change()
    Worksheets(Me.cboWhichSheet.Value).Sheet
End Sub

Private Sub cmdAddWorksheet_Click()
    Worksheets.Add before:=Worksheets(1)
    
    ActiveSheet.Name = InputBox("Enter new sheet name")
    Me.cboWhichSheet.AddItem (ActiveSheet.Name)
    
End Sub

Private Sub cmdCreateReport_Click()
    FinalReport
    
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer

    For i = 1 To Worksheets.Count
        Worksheets(i).Select
        Me.cboWhichSheet.AddItem (Worksheets(i).Name)
        
    Next i
End Sub
