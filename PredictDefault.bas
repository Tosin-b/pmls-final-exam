Attribute VB_Name = "PredictDefault"
' Reads the feature values from the Input sheet, posts them to the FastAPI app,
' and writes the returned probability of default back into the sheet.
'
' Input sheet layout: feature names in column A, values in column B, starting at row 2,
' in the same order as the "features" list saved in model.pkl.

Sub PredictDefault()

    ' change this one line to point at a deployed URL instead of the local app
    Const API_URL As String = "http://127.0.0.1:8000/predict"

    Dim ws As Worksheet
    Dim http As Object
    Dim json As String, response As String
    Dim name As String, value As String
    Dim row As Long
    Dim prob As Double
    Dim startPos As Long, endPos As Long

    Set ws = ThisWorkbook.Sheets("Input")

    ' build the JSON body from the name/value pairs in columns A and B
    json = "{"
    row = 2
    Do While ws.Cells(row, 1).Value <> ""
        name = Trim(ws.Cells(row, 1).Value)
        value = Trim(CStr(ws.Cells(row, 2).Value))

        If value = "" Then
            MsgBox "No value entered for " & name & " (cell B" & row & ").", vbExclamation
            Exit Sub
        End If

        ' make sure the decimal separator is a dot, not a comma
        value = Replace(value, ",", ".")

        If row > 2 Then json = json & ", "
        json = json & """" & name & """: " & value
        row = row + 1
    Loop
    json = json & "}"

    If row = 2 Then
        MsgBox "No features found on the Input sheet. Put names in column A and values in column B, starting at row 2.", vbExclamation
        Exit Sub
    End If

    ' send it
    On Error GoTo ApiError
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "POST", API_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send json
    On Error GoTo 0

    If http.Status <> 200 Then
        MsgBox "The API returned status " & http.Status & "." & vbCrLf & vbCrLf & _
               http.responseText, vbExclamation
        Exit Sub
    End If

    response = http.responseText

    ' pull the number out of {"probability_of_default": 0.2715} with plain string handling
    startPos = InStr(response, ":")
    endPos = InStr(response, "}")
    If startPos = 0 Or endPos = 0 Then
        MsgBox "Could not read the probability from the response:" & vbCrLf & response, vbExclamation
        Exit Sub
    End If

    prob = CDbl(Replace(Trim(Mid(response, startPos + 1, endPos - startPos - 1)), ".", _
                        Application.International(xlDecimalSeparator)))

    ' write it next to the "Probability of Default" label
    ws.Range("E2").Value = "Probability of Default"
    ws.Range("E3").Value = prob
    ws.Range("E3").NumberFormat = "0.00%"
    ws.Range("E2:E3").HorizontalAlignment = xlLeft

    MsgBox "Predicted probability of default: " & Format(prob, "0.00%"), vbInformation
    Exit Sub

ApiError:
    MsgBox "Could not reach the API at " & API_URL & "." & vbCrLf & vbCrLf & _
           "Start it first with:" & vbCrLf & _
           "    uvicorn main:app --reload" & vbCrLf & vbCrLf & _
           "then try again.", vbExclamation

End Sub
