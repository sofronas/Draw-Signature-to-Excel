Private Sub CommandButton1_Click()
    Dim oShell As Object, oCmd As String
    Dim oExec As Object, oOutput As Object
    Dim arg As Variant
    Dim s As String, sLine As String

    Set oShell = CreateObject("WScript.Shell")
    arg = "somevalue"
    oCmd = "pythonw ""C:\Users\user\Desktop\pop.pyw""" & " " & arg

    Set oExec = oShell.Exec(oCmd)
    Set oOutput = oExec.StdOut

    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbNewLine
    Wend

    Debug.Print s

    Set oOutput = Nothing: Set oExec = Nothing
    Set oShell = Nothing
    
    'Range("c3").ClearContents
    Dim s1 As String
    Dim pic As Picture
    Dim rng As Range
    
    For Each shp In ActiveSheet.Shapes
        If Not Intersect(shp.TopLeftCell, [C3:C4]) Is Nothing Then shp.Delete
        Next
    
    Image_Name = "img"
    Image_Location = "C:\Users\user\Desktop\"
    Image_Format = ".png"
    Cell_Reference = "C3"
    
    Set Image = ActiveSheet.Pictures.Insert(Image_Location + "\" + Image_Name + Image_Format)
    Image.Top = Range(Cell_Reference).Top
    Image.Left = Range(Cell_Reference).Left
    Image.ShapeRange.Height = 15
    Image.ShapeRange.Width = 105
End Sub
