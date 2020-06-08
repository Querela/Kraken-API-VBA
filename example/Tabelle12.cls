VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle12"
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
Private Const cellLinkButton As String = "B15"
Private Const cellOutputErrMsg As String = "D15"
Private Const offsetTopResults As Integer = 19
Private Const rangeResultsToClear As String = "B19:Z118"

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

Private Sub DoGetOpenOrders()
    Dim creds As Dictionary
    Dim method As String
    Dim params As New Dictionary
    Dim sTrades, sUserref As String
    Dim result As Variant, errJson As Variant
    Dim i As Integer, curOffset As Integer
    Dim arrKeys() As Variant, key As String, item As Variant, itemDescr As Variant

    Set creds = LoadCredentials()
    If creds Is Nothing Then
        ActiveSheet.Range(cellOutputErrMsg).Value = "No Key & Secret and no credentials file found!"
        Exit Sub
    End If

    ' prepare params
    method = "OpenOrders"

    ' Check if assetpair given
    sTrades = ActiveSheet.Range(cellInputTrades).Value
    If Len(sTrades) > 0 Then
        params.Add "trades", sTrades
    End If
    sUserref = ActiveSheet.Range(cellInputUserref).Value
    If Len(sUserref) > 0 Then
        params.Add "userref", sUserref
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
    Set result = result("open")
    arrKeys = result.Keys

    For i = 0 To UBound(arrKeys)
        If i Mod 10 = 0 Then VBA.DoEvents

        key = arrKeys(i)
        Set item = result(key)
        curOffset = offsetTopResults + i

        ActiveSheet.Range("B" & curOffset).Value = key
        ActiveSheet.Range("C" & curOffset).Value = item("refid")
        ActiveSheet.Range("D" & curOffset).Value = item("userref")
        ActiveSheet.Range("E" & curOffset).Value = item("status")
        ActiveSheet.Range("F" & curOffset).Value = item("opentm")
        ActiveSheet.Range("G" & curOffset).Value = item("starttm")
        ActiveSheet.Range("H" & curOffset).Value = item("expiretm")

        Set itemDescr = item("descr")
        ActiveSheet.Range("I" & curOffset).Value = itemDescr("pair")
        ActiveSheet.Range("J" & curOffset).Value = itemDescr("type")
        ActiveSheet.Range("K" & curOffset).Value = itemDescr("ordertype")
        ActiveSheet.Range("L" & curOffset).Value = itemDescr("price")
        ActiveSheet.Range("M" & curOffset).Value = itemDescr("price2")
        ActiveSheet.Range("N" & curOffset).Value = itemDescr("leverage")
        ActiveSheet.Range("O" & curOffset).Value = itemDescr("order")
        ActiveSheet.Range("P" & curOffset).Value = itemDescr("close")

        ActiveSheet.Range("Q" & curOffset).Value = item("vol")
        ActiveSheet.Range("R" & curOffset).Value = item("vol_exec")
        ActiveSheet.Range("S" & curOffset).Value = item("cost")
        ActiveSheet.Range("T" & curOffset).Value = item("fee")
        ActiveSheet.Range("U" & curOffset).Value = item("price")
        ActiveSheet.Range("V" & curOffset).Value = item("stopprice")
        ActiveSheet.Range("W" & curOffset).Value = item("limitprice")
        ActiveSheet.Range("X" & curOffset).Value = item("misc")
        ActiveSheet.Range("Y" & curOffset).Value = item("oflags")

        If item.Exists("trades") Then
            Set itemDescr = item("trades")
            ActiveSheet.Range("Z" & curOffset).Value = ""
        End If
    Next
End Sub


' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String
    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get Open Order ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetOpenOrders
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################
