Attribute VB_Name = "Main"

Sub getInfoFromColumn(nCol As Integer)
    
    JsonConverter.JsonOptions.AllowUnquotedKeys = True
    
    Dim ddObj As DaDataObject
    Dim lRow As Integer
    Dim i As Integer
    
    lRow = myLib.getLastRow
    For i = 1 To lRow
        Set ddObj = New DaDataObject
        ddObj.InitiateProperties JsonConverter.ParseJson(getResponse(Cells(i, nCol), "address", "f21c3e7d77083b1d0d18d3f6a4b0ee5c18e521a8"))
            Cells(i, nCol + 1) = ddObj.country
            Cells(i, nCol + 2) = ddObj.region_fias_id
            Cells(i, nCol + 3) = ddObj.region_kladr_id
            Cells(i, nCol + 4) = ddObj.region_with_type
            Cells(i, nCol + 5) = ddObj.region_type
            Cells(i, nCol + 6) = ddObj.region_type_full
            Cells(i, nCol + 7) = ddObj.region
            Cells(i, nCol + 8) = ddObj.area_fias_id
            Cells(i, nCol + 9) = ddObj.area_kladr_id
            Cells(i, nCol + 10) = ddObj.area_with_type
            Cells(i, nCol + 11) = ddObj.area_type
            Cells(i, nCol + 12) = ddObj.area_type_full
            Cells(i, nCol + 13) = ddObj.area
            Cells(i, nCol + 14) = ddObj.city_fias_id
            Cells(i, nCol + 15) = ddObj.city_kladr_id
            Cells(i, nCol + 16) = ddObj.city_with_type
            Cells(i, nCol + 17) = ddObj.city_type
            Cells(i, nCol + 18) = ddObj.city_type_full
            Cells(i, nCol + 19) = ddObj.city
            Cells(i, nCol + 20) = ddObj.city_area
            Cells(i, nCol + 21) = ddObj.city_district
            Cells(i, nCol + 22) = ddObj.settlement_fias_id
            Cells(i, nCol + 23) = ddObj.settlement_kladr_id
            Cells(i, nCol + 24) = ddObj.settlement_with_type
            Cells(i, nCol + 25) = ddObj.settlement_type
            Cells(i, nCol + 26) = ddObj.settlement_type_full
            Cells(i, nCol + 27) = ddObj.settlement
            Cells(i, nCol + 28) = ddObj.street_fias_id
            Cells(i, nCol + 29) = ddObj.street_kladr_id
            Cells(i, nCol + 30) = ddObj.street_with_type
            Cells(i, nCol + 31) = ddObj.street_type
            Cells(i, nCol + 32) = ddObj.street_type_full
            Cells(i, nCol + 33) = ddObj.street
            Cells(i, nCol + 34) = ddObj.house_fias_id
            Cells(i, nCol + 35) = ddObj.house_kladr_id
            Cells(i, nCol + 36) = ddObj.house_type
            Cells(i, nCol + 37) = ddObj.house_type_full
            Cells(i, nCol + 38) = ddObj.house
            Cells(i, nCol + 39) = ddObj.block_type
            Cells(i, nCol + 40) = ddObj.block_type_full
            Cells(i, nCol + 41) = ddObj.block
            Cells(i, nCol + 42) = ddObj.flat_type
            Cells(i, nCol + 43) = ddObj.flat_type_full
            Cells(i, nCol + 44) = ddObj.flat
            Cells(i, nCol + 45) = ddObj.flat_area
            Cells(i, nCol + 46) = ddObj.square_meter_price
            Cells(i, nCol + 47) = ddObj.flat_price
            Cells(i, nCol + 48) = ddObj.postal_box
            Cells(i, nCol + 49) = ddObj.fias_id
            Cells(i, nCol + 50) = ddObj.fias_level
            Cells(i, nCol + 51) = ddObj.kladr_id
            Cells(i, nCol + 52) = ddObj.capital_marker
            Cells(i, nCol + 53) = ddObj.okato
            Cells(i, nCol + 54) = ddObj.oktmo
            Cells(i, nCol + 55) = ddObj.tax_office
            Cells(i, nCol + 56) = ddObj.tax_office_legal
            Cells(i, nCol + 57) = ddObj.timezone
            Cells(i, nCol + 58) = ddObj.geo_lat
            Cells(i, nCol + 59) = ddObj.geo_lon
            Cells(i, nCol + 60) = ddObj.beltway_hit
            Cells(i, nCol + 61) = ddObj.beltway_distance
            Cells(i, nCol + 62) = ddObj.qc_geo
            Cells(i, nCol + 63) = ddObj.qc_complete
            Cells(i, nCol + 64) = ddObj.qc_house
            Cells(i, nCol + 65) = ddObj.qc
            Cells(i, nCol + 66) = ddObj.unparsed_parts
            Cells(i, nCol + 67) = ddObj.value
            Cells(i, nCol + 68) = ddObj.unrestricted_value
    Next i
End Sub

Sub test()
    getInfoFromColumn (1)
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
        .send ("{""query"":""" & Replace(Replace(text, Chr(34), ""), Chr(92), "") & """}")
    End With
    
    result = objHTTP.responseText
    result = Replace(result, "[", "")
    result = Replace(result, "]", "")
    
    If result = "{""suggestions"":}" Then result = "{""suggestions"":null}"
    
    getResponse = result
    
End Function
