VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputCredsKey As String = "D5"
Private Const cellInputCredsSecret As String = "D6"
Private Const cellInputCredsFilepath As String = "D8"
Private Const cellInputLedgerIDs As String = "D12"
Private Const cellLinkButton As String = "B14"
Private Const cellLinkClearBtn As String = "B15"
Private Const cellOutputErrMsg As String = "D14"
Private Const offsetTopResults As Integer = 18
Private Const rangeResultsToClear As String = "B18:J37"

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
    Dim sLedgerIDs As String

    sLedgerIDs = ActiveSheet.Range(cellInputLedgerIDs).Value
    If Len(sLedgerIDs) > 0 Then
        params.Add "id", sLedgerIDs
    End If

    Set GatherInputParameters = params
End Function

Private Sub DisplaySingleRow(ByVal key As String, ByRef item As Variant, ByVal iRow As Integer)
    ActiveSheet.Range("B" & iRow).Value = key
    ActiveSheet.Range("C" & iRow).Value = item("refid")
    ActiveSheet.Range("D" & iRow).Value = item("time")
    ActiveSheet.Range("E" & iRow).Value = item("type")
    ActiveSheet.Range("F" & iRow).Value = item("aclass")
    ActiveSheet.Range("G" & iRow).Value = item("asset")
    ActiveSheet.Range("H" & iRow).Value = item("amount")
    ActiveSheet.Range("I" & iRow).Value = item("fee")
    ActiveSheet.Range("J" & iRow).Value = item("balance")
End Sub

Private Sub DoGetQueryLedgers()
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
    method = "QueryLedgers"

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
        Debug.Print "Get Query Ledgers ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetQueryLedgers
        Application.ScreenUpdating = True
    ElseIf cellLinkClearBtn = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoClear
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################


