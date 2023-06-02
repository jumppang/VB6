Attribute VB_Name = "moMain"
Type nType
    aaa As Integer
    bbb As String
    
End Type

Global gCC As Integer


Sub main()
    Dim data As Integer
    Dim typeData As nType
    
    Call fileOpen
        
    gCC = 2
    typeData.aaa = 1
    
    Form1.Show
        
    
End Sub

Sub fileOpen()
    Dim fs, f
    Dim returnType
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile("d:\aaa\test.html")
    
    Set f = Nothing
    Set fs = Nothing
End Sub


Sub fileOpen2()
    Dim FineNum, Mode, Handle
    
    Open "D:\aaa\test.html" For Append As FileNum
    Mode = FileAttr(FileNum, 1)
    Handle = FileAttr(FineNum, 2) 'Return File Handle
    Close FileNum
    
    
    
End Sub

