Attribute VB_Name = "Main"
Sub test()
    JsonConverter.JsonOptions.AllowUnquotedKeys = True
    
    Dim json As Object
    Dim test2 As DaDataObject
    Dim response As String
    Dim address As String
    
    address = "Волгоградский проспект д. 125"
    response = getResponse(address, "address", "f21c3e7d77083b1d0d18d3f6a4b0ee5c18e521a8")
    
    Set json = JsonConverter.ParseJson(response)
    Set test2 = New DaDataObject
    
    On Error Resume Next
    test2.InitiateProperties json
    Cells(3, 10) = test2.house_kladr_id
    
End Sub

Function getResponse(ByVal text As String, request As String, key As String) As String

    Dim result As String
    Dim objHTTP As Object
    
    result = ""
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    With objHTTP
        URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/" & request
        .Open "POST", URL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Token " & key
        .setProxy 2, "128.114.0.21:8080", ""
        .send ("{""query"":""" & text & """}")
    End With
    
    result = objHTTP.responseText
    result = Replace(result, "[", "")
    result = Replace(result, "]", "")
    
    If result = "{""suggestions"":}" Then result = "{""suggestions"":null}"
    
    getResponse = result
    
End Function
