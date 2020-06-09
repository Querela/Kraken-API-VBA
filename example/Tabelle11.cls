VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputCredsKey As String = "D5"
Private Const cellInputCredsSecret As String = "D6"
Private Const cellInputCredsFilepath As String = "D8"
Private Const cellInputAssetClass As String = "D12"
Private Const cellInputAsset As String = "D13"
Private Const cellLinkButton As String = "B15"
Private Const cellLinkClearBtn As String = "B16"
Private Const cellOutputErrMsg As String = "D15"
Private Const offsetTopResults As Integer = 19
Private Const rangeResultsToClear As String = "B19:B26"  ' "B19:B27"

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
    Dim sAssetClass, sAsset As String

    sAssetClass = ActiveSheet.Range(cellInputAssetClass).Value
    If Len(sAssetClass) > 0 Then
        params.Add "aclass", sAssetClass
    End If
    sAsset = ActiveSheet.Range(cellInputAsset).Value
    If Len(sAsset) > 0 Then
        params.Add "asset", sAsset
    End If

    Set GatherInputParameters = params
End Function

Private Sub DoGetTradeBalance()
    Dim creds As Dictionary
    Dim method As String
    Dim params As New Dictionary
    Dim result As Variant, errJson As Variant
    'Dim i As Integer, curOffset As Integer
    'Dim arrKeys() As Variant, key As String, sValue As String

    Set creds = LoadCredentials()
    If creds Is Nothing Then
        ActiveSheet.Range(cellOutputErrMsg).Value = "No Key & Secret and no credentials file found!"
        Exit Sub
    End If

    ' prepare params
    method = "TradeBalance"

    ' gather input
    Set params = GatherInputParameters()

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
    'arrKeys = result.Keys

    'For i = 0 To UBound(arrKeys)
    '    If i Mod 10 = 0 Then VBA.DoEvents

    '    key = arrKeys(i)
    '    sValue = result(key)
    '    curOffset = offsetTopResults + i

    '    ActiveSheet.Range("B" & curOffset).Value = sValue
    'Next

    ActiveSheet.Range("B" & (offsetTopResults + 0)).Value = result("eb")
    ActiveSheet.Range("B" & (offsetTopResults + 1)).Value = result("tb")
    ActiveSheet.Range("B" & (offsetTopResults + 2)).Value = result("m")
    ActiveSheet.Range("B" & (offsetTopResults + 3)).Value = result("n")
    ActiveSheet.Range("B" & (offsetTopResults + 4)).Value = result("c")
    ActiveSheet.Range("B" & (offsetTopResults + 5)).Value = result("v")
    ActiveSheet.Range("B" & (offsetTopResults + 6)).Value = result("e")
    ActiveSheet.Range("B" & (offsetTopResults + 7)).Value = result("mf")
    ' seems to be empty?
    'ActiveSheet.Range("B" & (offsetTopResults + 8)).Value = result("ml")
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
        Debug.Print "Get Trade Balance ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetTradeBalance
        Application.ScreenUpdating = True
    ElseIf cellLinkClearBtn = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoClear
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################
