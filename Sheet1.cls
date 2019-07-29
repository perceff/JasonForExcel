VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Option Explicit
'-----------------------------------------------------------------
Private Const API_KEY = "Введите свой ключ"
' 2. запускается вторым
Function Suggest(ByVal Name, ByVal query, ByVal Count) As String
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    timeout = 2000 'milliseconds
    http.setTimeouts timeout, timeout, timeout, timeout
    request = "{ ""query"": """ & query & """, ""count"": """ & Count & """ }"
    http.Open "POST", "https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/" & Name
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Token " & API_KEY
    http.Send request
    suggestions = Replace(http.responseText, "\""", "*") 'Replace(http.responseText, "\""", "")
    suggestions = Replace(suggestions, Chr(10), "")
    Debug.Print suggestions
    Suggest = suggestions
End Function

' 1. запускается первым
Public Sub FillTable()
Msg = MsgBox("Начать заполнение?", 1, "Предупреждение")
Application.ScreenUpdating = False
If Msg = 2 Then Exit Sub
    On Error Resume Next
    Dim Json As Object
    r = 3: c = 6
    Do While ThisWorkbook.Worksheets(1).Range(Cells(r, c), Cells(r, c)).Value <> ""
        suggested = Suggest("party", ThisWorkbook.Worksheets(1).Range(Cells(r, c), Cells(r, c)).Value, 1)
        If Len(suggested) < 20 Then GoTo ER
        JsonConverter.JsonOptions.AllowUnquotedKeys = True
        JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
        JsonConverter.JsonOptions.EscapeSolidus = True
        Set Json = JsonConverter.ParseJson(suggested)
        With ThisWorkbook.Worksheets(1)
            .Cells(r, c + 1).Value = Json("suggestions")(1)("data")("kpp")
            .Cells(r, c + 2).Value = Json("suggestions")(1)("data")("ogrn")
            .Cells(r, c + 3).Value = Json("suggestions")(1)("data")("okved")
            .Cells(r, c + 4).Value = Replace(Json("suggestions")(1)("data")("name")("full_with_opf"), "*", """")
            .Cells(r, c + 5).Value = Replace(Json("suggestions")(1)("data")("name")("short_with_opf"), "*", """")
            .Cells(r, c + 6).Value = Replace(Json("suggestions")(1)("data")("name")("short"), "*", """")
            .Cells(r, c + 7).Value = Json("suggestions")(1)("data")("opf")("short")
            .Cells(r, c + 8).Value = Json("suggestions")(1)("data")("address")("data")("fias_id")
            .Cells(r, c + 9).Value = Json("suggestions")(1)("data")("address")("data")("fias_level")
            .Cells(r, c + 10).Value = Json("suggestions")(1)("data")("postal_code")
            .Cells(r, c + 11).Value = Json("suggestions")(1)("data")("address")("data")("region_with_type") & " " & Json("suggestions")(1)("data")("address")("data")("area_with_type")
            If Json("suggestions")(1)("data")("address")("data")("city_with_type") = "" Then
                    .Cells(r, c + 12).Value = Json("suggestions")(1)("data")("address")("data")("city_with_type")
                Else
                    .Cells(r, c + 12).Value = Json("suggestions")(1)("data")("address")("data")("settlement_with_type")
            End If
            .Cells(r, c + 13).Value = Json("suggestions")(1)("data")("address")("data")("city_district_with_type")
            .Cells(r, c + 14).Value = Json("suggestions")(1)("data")("address")("data")("street_with_type")
            .Cells(r, c + 15).Value = Json("suggestions")(1)("data")("address")("data")("house")
            .Cells(r, c + 16).Value = Json("suggestions")(1)("data")("address")("data")("block_type") & " " & Json("suggestions")(1)("data")("address")("data")("block")
            .Cells(r, c + 17).Value = Json("suggestions")(1)("data")("address")("data")("flat")
            .Cells(r, c + 18).Value = Json("suggestions")(1)("data")("address")("value")
            .Cells(r, c + 19).Value = Json("suggestions")(1)("data")("address")("data")("source")
            .Cells(r, c + 20).Value = Json("suggestions")(1)("data")("management")("name")
            .Cells(r, c + 21).Value = Json("suggestions")(1)("data")("management")("post")
            .Cells(r, c + 22).Value = Json("suggestions")(1)("data")("state")("status")
            .Cells(r, c + 23).Value = Json("suggestions")(1)("data")("state")("actuality_date")
        End With
ER:     r = r + 1
    suggested = "": suggestions = ""
    Set Json = Nothing
    Sleep (50)
    Loop
Application.ScreenUpdating = True
End Sub
'-------------------код для проверки единичной записи--------------------------------
'Private Sub Worksheet_Change(ByVal target As Range)
'Msg = MsgBox("Начать заполнение?", 1, "Предупреждение")
'If Msg = 2 Then Exit Sub
'    On Error Resume Next
'    Dim Json As Object
'    r = 3: c = 6
'        suggested = Suggest("party", ThisWorkbook.Worksheets(1).Range(Cells(r, c), Cells(r, c)).Value, 1)
'        If Len(suggested) < 20 Then GoTo ER
'        JsonConverter.JsonOptions.AllowUnquotedKeys = True
'        JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
'        JsonConverter.JsonOptions.EscapeSolidus = True
'        Set Json = JsonConverter.ParseJson(suggested)
'        With ThisWorkbook.Worksheets(1)
'            .Cells(r, c + 1).Value = Json("suggestions")(1)("data")("kpp")
'            .Cells(r, c + 2).Value = Json("suggestions")(1)("data")("ogrn")
'            .Cells(r, c + 3).Value = Json("suggestions")(1)("data")("okved")
'            .Cells(r, c + 4).Value = Replace(Json("suggestions")(1)("data")("name")("full_with_opf"), "*", """")
'            .Cells(r, c + 5).Value = Replace(Json("suggestions")(1)("data")("name")("short_with_opf"), "*", """")
'            .Cells(r, c + 6).Value = Replace(Json("suggestions")(1)("data")("name")("short"), "*", """")
'            .Cells(r, c + 7).Value = Json("suggestions")(1)("data")("opf")("short")
'            .Cells(r, c + 8).Value = Json("suggestions")(1)("data")("address")("data")("fias_id")
'            .Cells(r, c + 9).Value = Json("suggestions")(1)("data")("address")("data")("fias_level")
'            .Cells(r, c + 10).Value = Json("suggestions")(1)("data")("postal_code")
'            .Cells(r, c + 11).Value = Json("suggestions")(1)("data")("address")("data")("region_with_type") & " " & Json("suggestions")(1)("data")("address")("data")("area_with_type")
'            If Json("suggestions")(1)("data")("address")("data")("city_with_type") = "" Then
'                    .Cells(r, c + 12).Value = Json("suggestions")(1)("data")("address")("data")("city_with_type")
'                Else
'                    .Cells(r, c + 12).Value = Json("suggestions")(1)("data")("address")("data")("settlement_with_type")
'            End If
'            .Cells(r, c + 13).Value = Json("suggestions")(1)("data")("address")("data")("city_district_with_type")
'            .Cells(r, c + 14).Value = Json("suggestions")(1)("data")("address")("data")("street_with_type")
'            .Cells(r, c + 15).Value = Json("suggestions")(1)("data")("address")("data")("house")
'            .Cells(r, c + 16).Value = Json("suggestions")(1)("data")("address")("data")("block_type") & " " & Json("suggestions")(1)("data")("address")("data")("block")
'            .Cells(r, c + 17).Value = Json("suggestions")(1)("data")("address")("data")("flat")
'            .Cells(r, c + 18).Value = Json("suggestions")(1)("data")("address")("value")
'            .Cells(r, c + 19).Value = Json("suggestions")(1)("data")("address")("data")("source")
'            .Cells(r, c + 20).Value = Json("suggestions")(1)("data")("management")("name")
'            .Cells(r, c + 21).Value = Json("suggestions")(1)("data")("management")("post")
'            .Cells(r, c + 22).Value = Json("suggestions")(1)("data")("state")("status")
'            .Cells(r, c + 23).Value = Json("suggestions")(1)("data")("state")("actuality_date")
'        End With
'ER:
'End Sub

