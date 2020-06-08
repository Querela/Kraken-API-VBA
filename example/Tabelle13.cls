VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputCredsKey As String = "D5"
Private Const cellInputCredsSecret As String = "D6"
Private Const cellInputCredsFilepath As String = "D8"
Private Const cellInputTrades As String = "D12"
Private Const cellInputUserref As String = "D13"
Private Const cellInputTimeStart As String = "D14"
Private Const cellInputTimeEnd As String = "D15"
Private Const cellInputResultOffset As String = "D16"
Private Const cellInputTimeType As String = "D17"
Private Const cellLinkButton As String = "B19"
Private Const cellOutputErrMsg As String = "D19"
Private Const offsetTopResults As Integer = 23
Private Const rangeResultsToClear As String = "B23:AB222"

' #############################################################################

Private Function LoadCredentials() As Dictionary
    Dim creds As Dictionary
    Dim filename, sKey, sSecret As String

    sKey = ActiveSheet.Range(cellInputCredsKey).Value
    sSecret = ActiveSheet.Range(cellInputCredsSecret).Value

    If Not ExcelUtils.IsStringEmpty(sKey) _
        And Not ExcelUtils.IsStringEmpty(sSecret) Then

        Set creds = New Dictionary
        creds.Add "key", Trim(sKey)
        creds.Add "secret", Trim(sSecret)

        Set LoadCredentials = creds
        Exit Function
    End If

    filename = ActiveSheet.Range(cellInputCredsFilepath).Value

    If ExcelUtils.IsStringEmpty(filename) Then
        filename = FileUtils.GetDefaultKrakenCredentialsFilepath()
    End If

    If FileUtils.ExistsKrakenCredentialsFile(filename) Then
        Set creds = FileUtils.LoadKrakenCredentials(filename)
    End If

    Set LoadCredentials = creds
End Function

Private Sub DoGetClosedOrders()
    Dim creds As Dictionary
    Dim method As String
    Dim params As New Dictionary
    Dim sTrades, sUserref As String
    Dim vTimeStart As Variant, vTimeEnd As Variant, sTimeType As String, iResultOffset As Integer
    Dim result As Variant, errJson As Variant
    Dim i As Integer, curOffset As Integer
    Dim arrKeys() As Variant, key As String, item As Variant, itemDescr As Variant
    Dim numResultsTotal, numResultsReturned, numResultsOffset As Integer

    Set creds = LoadCredentials()
    If creds Is Nothing Then
        ActiveSheet.Range(cellOutputErrMsg).Value = "No Key & Secret and no credentials file found!"
        Exit Sub
    End If

    ' prepare params
    method = "ClosedOrders"

    ' Check if assetpair given
    sTrades = ActiveSheet.Range(cellInputTrades).Value
    If Len(sTrades) > 0 Then
        params.Add "trades", sTrades
    End If
    sUserref = ActiveSheet.Range(cellInputUserref).Value
    If Len(sUserref) > 0 Then
        params.Add "userref", sUserref
    End If
    vTimeStart = ActiveSheet.Range(cellInputTimeStart).Value
    If Len(vTimeStart) > 0 Then
        params.Add "start", vTimeStart
    End If
    vTimeEnd = ActiveSheet.Range(cellInputTimeEnd).Value
    If Len(vTimeEnd) > 0 Then
        params.Add "end", vTimeEnd
    End If
    sTimeType = ActiveSheet.Range(cellInputTimeType).Value
    If Len(sTimeType) > 0 Then
        params.Add "closetime", sTimeType
    End If
    iResultOffset = ActiveSheet.Range(cellInputResultOffset).Value
    If Len(iResultOffset) > 0 Then
        ' TODO: hmm, check is weird ...
        params.Add "ofs", iResultOffset
    End If

    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPrivate(creds("key"), creds("secret"), method, params)
    Debug.Print method, WebUtils.BeautifyJson(result)

    errJson = result("error")
    If Not IsEmpty(errJson) Then
        ActiveSheet.Range(cellOutputErrMsg).Value = errJson(0)
        Exit Sub
    Else
        ActiveSheet.Range(cellOutputErrMsg).ClearContents
    End If

    ' TODO: Q'n'D clearing ...
    ActiveSheet.Range(rangeResultsToClear).ClearContents

    Set result = result("result")

    numResultsTotal = result("count")
    Debug.Print "Number of results:", numResultsTotal

    Set result = result("closed")
    arrKeys = result.Keys

    ' just shove new offset into input mask ...
    numResultsReturned = iResultOffset + UBound(arrKeys) + 1
    ActiveSheet.Range(cellInputResultOffset).Value = numResultsReturned

    For i = 0 To UBound(arrKeys)
        If i Mod 10 = 0 Then VBA.DoEvents

        key = arrKeys(i)
        Set item = result(key)
        curOffset = offsetTopResults + i

        ActiveSheet.Range("B" & curOffset).Value = key
        ActiveSheet.Range("C" & curOffset).Value = item("refid")
        ActiveSheet.Range("D" & curOffset).Value = item("userref")
        ActiveSheet.Range("E" & curOffset).Value = item("status")
        ActiveSheet.Range("F" & curOffset).Value = item("reason")
        ActiveSheet.Range("G" & curOffset).Value = item("opentm")
        ActiveSheet.Range("H" & curOffset).Value = item("closetm")
        ActiveSheet.Range("I" & curOffset).Value = item("starttm")
        ActiveSheet.Range("J" & curOffset).Value = item("expiretm")

        Set itemDescr = item("descr")
        ActiveSheet.Range("K" & curOffset).Value = itemDescr("pair")
        ActiveSheet.Range("L" & curOffset).Value = itemDescr("type")
        ActiveSheet.Range("M" & curOffset).Value = itemDescr("ordertype")
        ActiveSheet.Range("N" & curOffset).Value = itemDescr("price")
        ActiveSheet.Range("O" & curOffset).Value = itemDescr("price2")
        ActiveSheet.Range("P" & curOffset).Value = itemDescr("leverage")
        ActiveSheet.Range("Q" & curOffset).Value = itemDescr("order")
        ActiveSheet.Range("R" & curOffset).Value = itemDescr("close")

        ActiveSheet.Range("S" & curOffset).Value = item("vol")
        ActiveSheet.Range("T" & curOffset).Value = item("vol_exec")
        ActiveSheet.Range("U" & curOffset).Value = item("cost")
        ActiveSheet.Range("V" & curOffset).Value = item("fee")
        ActiveSheet.Range("W" & curOffset).Value = item("price")
        ActiveSheet.Range("X" & curOffset).Value = item("stopprice")
        ActiveSheet.Range("Y" & curOffset).Value = item("limitprice")
        ActiveSheet.Range("Z" & curOffset).Value = item("misc")
        ActiveSheet.Range("AA" & curOffset).Value = item("oflags")

        If item.Exists("trades") Then
            Set itemDescr = item("trades")
            ActiveSheet.Range("AB" & curOffset).Value = ""
        End If
    Next
End Sub


' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String
    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get Closed Order ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetClosedOrders
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################


