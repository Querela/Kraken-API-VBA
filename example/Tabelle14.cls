VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle14"
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
Private Const cellInputTXIDs As String = "D14"
Private Const cellLinkButton As String = "B16"
Private Const cellLinkClearBtn As String = "B17"
Private Const cellOutputErrMsg As String = "D16"
Private Const offsetTopResults As Integer = 20
Private Const rangeResultsToClear As String = "B20:AB69"

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

Private Function GatherInputParameters() As Dictionary
    Dim params As New Dictionary
    Dim sTrades, sUserref As String, sTXIDs As String

    sTrades = ActiveSheet.Range(cellInputTrades).Value
    If Len(sTrades) > 0 Then
        params.Add "trades", sTrades
    End If
    sUserref = ActiveSheet.Range(cellInputUserref).Value
    If Len(sUserref) > 0 Then
        params.Add "userref", sUserref
    End If
    sTXIDs = ActiveSheet.Range(cellInputTXIDs).Value
    If Len(sTXIDs) > 0 Then
        params.Add "txid", sTXIDs
    End If

    Set GatherInputParameters = params
End Function

Private Sub DisplaySingleRow(ByVal key As String, ByRef item As Variant, ByVal iRow As Integer)
    Dim itemDescr As Variant

    ActiveSheet.Range("B" & iRow).Value = key
    ActiveSheet.Range("C" & iRow).Value = item("refid")
    ActiveSheet.Range("D" & iRow).Value = item("userref")
    ActiveSheet.Range("E" & iRow).Value = item("status")
    ActiveSheet.Range("F" & iRow).Value = item("reason")
    ActiveSheet.Range("G" & iRow).Value = item("opentm")
    ActiveSheet.Range("H" & iRow).Value = item("closetm")
    ActiveSheet.Range("I" & iRow).Value = item("starttm")
    ActiveSheet.Range("J" & iRow).Value = item("expiretm")

    Set itemDescr = item("descr")
    ActiveSheet.Range("K" & iRow).Value = itemDescr("pair")
    ActiveSheet.Range("L" & iRow).Value = itemDescr("type")
    ActiveSheet.Range("M" & iRow).Value = itemDescr("ordertype")
    ActiveSheet.Range("N" & iRow).Value = itemDescr("price")
    ActiveSheet.Range("O" & iRow).Value = itemDescr("price2")
    ActiveSheet.Range("P" & iRow).Value = itemDescr("leverage")
    ActiveSheet.Range("Q" & iRow).Value = itemDescr("order")
    ActiveSheet.Range("R" & iRow).Value = itemDescr("close")

    ActiveSheet.Range("S" & iRow).Value = item("vol")
    ActiveSheet.Range("T" & iRow).Value = item("vol_exec")
    ActiveSheet.Range("U" & iRow).Value = item("cost")
    ActiveSheet.Range("V" & iRow).Value = item("fee")
    ActiveSheet.Range("W" & iRow).Value = item("price")
    ActiveSheet.Range("X" & iRow).Value = item("stopprice")
    ActiveSheet.Range("Y" & iRow).Value = item("limitprice")
    ActiveSheet.Range("Z" & iRow).Value = item("misc")
    ActiveSheet.Range("AA" & iRow).Value = item("oflags")

    If item.Exists("trades") Then
        Set itemDescr = item("trades")
        ActiveSheet.Range("AB" & iRow).Value = ""
    End If
End Sub

Private Sub DoGetQueryOrdersInfo()
    Dim creds As Dictionary
    Dim method As String
    Dim params As Dictionary
    Dim iResultOffset As Integer
    Dim result As Variant, errJson As Variant
    Dim i As Integer, curOffset As Integer
    Dim arrKeys() As Variant, key As String, item As Variant
    Dim numResultsTotal, numResultsReturned, numResultsOffset As Integer

    ' load credentials
    Set creds = LoadCredentials()
    If creds Is Nothing Then
        ActiveSheet.Range(cellOutputErrMsg).Value = "No Key & Secret and no credentials file found!"
        Exit Sub
    End If

    ' prepare params
    method = "QueryOrders"

    ' gather inputs
    Set params = GatherInputParameters()

    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPrivate(creds("key"), creds("secret"), method, params)
    Debug.Print method, WebUtils.BeautifyJson(result)

    ' check and handle errors
    errJson = result("error")
    If Not IsEmpty(errJson) Then
        ActiveSheet.Range(cellOutputErrMsg).Value = errJson(0)
        Exit Sub
    Else
        ActiveSheet.Range(cellOutputErrMsg).ClearContents
    End If

    ' TODO: Q'n'D clearing ...
    ActiveSheet.Range(rangeResultsToClear).ClearContents

    ' process results
    Set result = result("result")
    arrKeys = result.Keys

    ' output rows
    For i = 0 To UBound(arrKeys)
        If i Mod 10 = 0 Then VBA.DoEvents

        key = arrKeys(i)
        Set item = result(key)
        curOffset = offsetTopResults + i

        DisplaySingleRow key, item, curOffset
    Next
End Sub

Private Sub DoClear()
    ' TODO: Q'n'D clearing ...
    ActiveSheet.Range(rangeResultsToClear).ClearContents
End Sub

' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String
    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get Query OrdersInfo ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetQueryOrdersInfo
        Application.ScreenUpdating = True
    ElseIf cellLinkClearBtn = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoClear
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################

