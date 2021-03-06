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
Private Const cellLinkClearBtn As String = "B16"
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

Private Function GatherInputParameters() As Dictionary
    Dim params As New Dictionary
    Dim sTrades, sUserref As String

    sTrades = ActiveSheet.Range(cellInputTrades).Value
    If Len(sTrades) > 0 Then
        params.Add "trades", sTrades
    End If
    sUserref = ActiveSheet.Range(cellInputUserref).Value
    If Len(sUserref) > 0 Then
        params.Add "userref", sUserref
    End If

    Set GatherInputParameters = params
End Function

Private Sub DisplaySingleRow(ByVal key As String, ByRef item As Variant, ByVal iRow As Integer)
    Dim itemDescr As Variant

    ActiveSheet.Range("B" & iRow).Value = key
    ActiveSheet.Range("C" & iRow).Value = item("refid")
    ActiveSheet.Range("D" & iRow).Value = item("userref")
    ActiveSheet.Range("E" & iRow).Value = item("status")
    ActiveSheet.Range("F" & iRow).Value = item("opentm")
    ActiveSheet.Range("G" & iRow).Value = item("starttm")
    ActiveSheet.Range("H" & iRow).Value = item("expiretm")

    Set itemDescr = item("descr")
    ActiveSheet.Range("I" & iRow).Value = itemDescr("pair")
    ActiveSheet.Range("J" & iRow).Value = itemDescr("type")
    ActiveSheet.Range("K" & iRow).Value = itemDescr("ordertype")
    ActiveSheet.Range("L" & iRow).Value = itemDescr("price")
    ActiveSheet.Range("M" & iRow).Value = itemDescr("price2")
    ActiveSheet.Range("N" & iRow).Value = itemDescr("leverage")
    ActiveSheet.Range("O" & iRow).Value = itemDescr("order")
    ActiveSheet.Range("P" & iRow).Value = itemDescr("close")

    ActiveSheet.Range("Q" & iRow).Value = item("vol")
    ActiveSheet.Range("R" & iRow).Value = item("vol_exec")
    ActiveSheet.Range("S" & iRow).Value = item("cost")
    ActiveSheet.Range("T" & iRow).Value = item("fee")
    ActiveSheet.Range("U" & iRow).Value = item("price")
    ActiveSheet.Range("V" & iRow).Value = item("stopprice")
    ActiveSheet.Range("W" & iRow).Value = item("limitprice")
    ActiveSheet.Range("X" & iRow).Value = item("misc")
    ActiveSheet.Range("Y" & iRow).Value = item("oflags")

    If item.Exists("trades") Then
        Set itemDescr = item("trades")
        ActiveSheet.Range("Z" & iRow).Value = ""
    End If
End Sub

Private Sub DoGetOpenOrders()
    Dim creds As Dictionary
    Dim method As String
    Dim params As Dictionary
    Dim result As Variant, errJson As Variant
    Dim i As Integer, curOffset As Integer
    Dim arrKeys() As Variant, key As String, item As Variant

    ' load credentials
    Set creds = LoadCredentials()
    If creds Is Nothing Then
        ActiveSheet.Range(cellOutputErrMsg).Value = "No Key & Secret and no credentials file found!"
        Exit Sub
    End If

    ' prepare params
    method = "OpenOrders"

    ' gather input
    Set params = GatherInputParameters()

    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPrivate(creds("key"), creds("secret"), method, params)
    Debug.Print method, WebUtils.BeautifyJson(result)

    ' check & handle errors
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
        Debug.Print "Get Open Order ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetOpenOrders
        Application.ScreenUpdating = True
    ElseIf cellLinkClearBtn = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoClear
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################
