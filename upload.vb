Sub CheckStatus(my_arr As Variant)
    
    On Error Resume Next
    
    Dim pd                      As String
    Dim mu                      As String
    Dim wi                      As Object
    
    Set wi = CreateObject("WinHttp.WinHttpRequest.5.1")

    mu = "https://docs.google.com/forms/d//formResponse"
    
    pd = "entry_479868114=" & my_arr(0) & _
                "&entry_1155996727=" & my_arr(1) & _
                "&entry_922606695=" & my_arr(2) & _
                "&entry_1990943469=" & my_arr(3)
    
    wi.Open "POST", mu, False
    wi.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    wi.Send (pd)

    'result = winHttpReq.responseText
    On Error GoTo 0

End Sub
