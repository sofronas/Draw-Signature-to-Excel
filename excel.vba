
Private Sub CommandButton1_Click()
    Dim oShell As Object, oCmd As String
    Dim oExec As Object, oOutput As Object
    Dim arg As Variant
    Dim s As String, sLine As String

    Set oShell = CreateObject("WScript.Shell")
    arg = "somevalue"
        
    'Python Version
    oCmd = "pythonw ""C:\Users\support\Desktop\pop.pyw""" & " " & arg
    Set oExec = oShell.Exec(oCmd)
    Set oOutput = oExec.StdOut
        
        
    'C# Version
    strProgramName = "C:\Users\user\WinFormsApp1\WinFormsApp1.exe"
    strArgument = "/G"
    
    
    Set oExec = oShell.Exec(strProgramName)
    Set oOutput = oExec.StdOut
    
    While Not oOutput.AtEndOfStream
        'sLine = oOutput.ReadLine
        'If sLine <> "" Then s = s & sLine & vbNewLine
    Wend

    'Debug.Print s

    Set oOutput = Nothing: Set oExec = Nothing
    Set oShell = Nothing
    
    'Range("c3").ClearContents
    
    'Delete signatures if exist
    Dim s1 As String
    Dim pic As Picture
    Dim rng As Range
    
    For Each shp In ActiveSheet.Shapes
        If Not Intersect(shp.TopLeftCell, [B45:B46]) Is Nothing Then shp.Delete
        Next
    
    For Each shp In ActiveSheet.Shapes
        If Not Intersect(shp.TopLeftCell, [L47:L48]) Is Nothing Then shp.Delete
        Next
    
    'Add First Signature
    Image_Name = "imgc"
    Image_Location = "C:\Users\support\Desktop\"
    Image_Format = ".png"
    Cell_Reference = "B45"
    
    Set Image = ActiveSheet.Pictures.Insert(Image_Location + "\" + Image_Name + Image_Format)
    Image.Top = Range(Cell_Reference).Top
    Image.Left = Range(Cell_Reference).Left
    Image.ShapeRange.LockAspectRatio = msoFalse
    Image.Placement = xlMoveAndSize
    Image.ShapeRange.Width = 144
    Image.ShapeRange.Height = 30
    
    'Add Second Signature
    Cell_Reference = "L47"
    Set Image = ActiveSheet.Pictures.Insert(Image_Location + "\" + Image_Name + Image_Format)
    Image.Top = Range(Cell_Reference).Top
    Image.Left = Range(Cell_Reference).Left
    Image.ShapeRange.LockAspectRatio = msoFalse
    Image.Placement = xlMoveAndSize
    Image.ShapeRange.Width = 144
    Image.ShapeRange.Height = 30
    
    'Add date
    Range("B48").Value = Date
    Range("B48").NumberFormat = "dd/mm/yyy"
    Range("L50").Value = Date
    Range("L50").NumberFormat = "dd/mm/yyy"
End Sub

