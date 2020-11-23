Attribute VB_Name = "Module1"
'
'Open Multiple Files
'
Public Sub ImportTextFile()
    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i As Integer
    
    OpenFiles = GetFiles()
    
    Application.ScreenUpdating = False
    
    For i = 1 To Application.CountA(OpenFiles)
    
        Set TextFile = Workbooks.Open(OpenFiles(i))
        
        TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
        Workbooks(1).Activate
        Workbooks(1).Worksheets.Add
        ActiveSheet.Paste
        ActiveSheet.Name = TextFile.Name
                
        Application.CutCopyMode = False
                
        TextFile.Close
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

Public Function GetFiles() As Variant
     GetFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
End Function
