VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputCredsKey As String = "D5"
Private Const cellInputCredsSecret As String = "D6"
Private Const cellInputCredsFilepath As String = "D8"
Private Const cellInputTypes As String = "D12"
Private Const cellInputTrades As String = "D13"
Private Const cellInputTimeStart As String = "D14"
Private Const cellInputTimeEnd As String = "D15"
Private Const cellInputResultOffset As String = "D16"
Private Const cellOutputNumResultsTotal As String = "E19"
Private Const cellLinkButton As String = "B18"
Private Const cellLinkClearBtn As String = "B19"
Private Const cellOutputErrMsg As String = "D18"
Private Const offsetTopResults As Integer = 22
Private Const rangeResultsToClear As String = "B22:V121"

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
    Dim sType, sTrades As String
    Dim vTimeStart As Variant, vTimeEnd As Variant, iResultOffset As Integer

    sType = ActiveSheet.Range(cellInputTypes).Value
    If Len(sType) > 0 Then
        params.Add "type", sType
    End If
    sTrades = ActiveSheet.Range(cellInputTrades).Value
    If Len(sTrades) > 0 Then
        params.Add "trades", sTrades
    End If
    vTimeStart = ActiveSheet.Range(cellInputTimeStart).Value
    If Len(vTimeStart) > 0 Then
        params.Add "start", vTimeStart
    End If
    vTimeEnd = ActiveSheet.Range(cellInputTimeEnd).Value
    If Len(vTimeEnd) > 0 Then
        params.Add "end", vTimeEnd
    End If
    iResultOffset = ActiveSheet.Range(cellInputResultOffset).Value
    If Len(iResultOffset) > 0 Then
        ' TODO: hmm, check is weird ...
        params.Add "ofs", iResultOffset
    End If

    Set GatherInputParameters = params
End Function

Private Sub DisplaySingleRow(ByVal key As String, ByRef item As Variant, ByVal iRow As Integer)
    ActiveSheet.Range("B" & iRow).Value = key
    ActiveSheet.Range("C" & iRow).Value = item("ordertxid")
    ActiveSheet.Range("D" & iRow).Value = item("postxid")
    ActiveSheet.Range("E" & iRow).Value = item("pair")
    ActiveSheet.Range("F" & iRow).Value = item("time")
    ActiveSheet.Range("G" & iRow).Value = item("type")
    ActiveSheet.Range("H" & iRow).Value = item("ordertype")
    ActiveSheet.Range("I" & iRow).Value = item("price")
    ActiveSheet.Range("J" & iRow).Value = item("cost")
    ActiveSheet.Range("K" & iRow).Value = item("fee")
    ActiveSheet.Range("L" & iRow).Value = item("vol")
    ActiveSheet.Range("M" & iRow).Value = item("margin")
    ActiveSheet.Range("N" & iRow).Value = item("misc")

    ActiveSheet.Range("O" & iRow).Value = item("posstatus")
    ActiveSheet.Range("P" & iRow).Value = item("cprice")
    ActiveSheet.Range("Q" & iRow).Value = item("ccost")
    ActiveSheet.Range("R" & iRow).Value = item("cfee")
    ActiveSheet.Range("S" & iRow).Value = item("cvol")
    ActiveSheet.Range("T" & iRow).Value = item("cmargin")
    ActiveSheet.Range("U" & iRow).Value = item("net")
    ActiveSheet.Range("V" & iRow).Value = item("trades")
End Sub

Private Sub DoGetTradesHistory()
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
    method = "TradesHistory"

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

    numResultsTotal = result("count")
    Debug.Print "Number of results:", numResultsTotal
    ActiveSheet.Range(cellOutputNumResultsTotal).Value = numResultsTotal

    Set result = result("trades")
    arrKeys = result.Keys

    ' just shove new offset into input mask ...
    iResultOffset = ActiveSheet.Range(cellInputResultOffset).Value
    numResultsReturned = iResultOffset + UBound(arrKeys) + 1
    ActiveSheet.Range(cellInputResultOffset).Value = numResultsReturned

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
    ActiveSheet.Range(cellOutputNumResultsTotal).ClearContents
End Sub

' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String
    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get Trades History ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetTradesHistory
        Application.ScreenUpdating = True
    ElseIf cellLinkClearBtn = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoClear
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################
