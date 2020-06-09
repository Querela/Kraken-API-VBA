VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputCredsKey As String = "D5"
Private Const cellInputCredsSecret As String = "D6"
Private Const cellInputCredsFilepath As String = "D8"
Private Const cellInputAssetPairs As String = "D12"
Private Const cellInputFeeInfos As String = "D13"
Private Const cellOutputCurrency As String = "C18"
Private Const cellOutputVolume As String = "C19"
Private Const cellLinkButton As String = "B15"
Private Const cellLinkClearBtn As String = "B16"
Private Const cellOutputErrMsg As String = "D15"
Private Const offsetTopResults As Integer = 23
Private Const rangeResultsToClear As String = "B23:N222" ' mhh, 23:181

' #############################################################################
' same as Tabelle5 (Kraken Ticker)

Private Function GetAllAssetPairs(Optional ByVal includeDarkpool As Boolean = False) As String
    Dim method As String
    Dim result As Variant, errJson As Variant
    Dim i As Integer, retVal As String, arrKeys() As Variant, key As String, item As Variant

    ' prepare params
    method = "AssetPairs"
    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPublic(method)
    'Debug.Print method, WebUtils.BeautifyJson(result)

    ' check error
    errJson = result("error")
    If Not IsEmpty(errJson) Then
        GetAllAssetPairs = Empty
        Exit Function
    End If

    ' concat asset pairs
    retVal = ""
    Set result = result("result")
    arrKeys = result.Keys
    For i = 0 To UBound(arrKeys)
        key = arrKeys(i)
        Set item = result(key)

        If Not includeDarkpool And ExcelUtils.EndsWith(key, ".d") Then GoTo NextIteration

        If i > 0 Then
            retVal = retVal & ","
        End If

        retVal = retVal & item("altname")
NextIteration:
    Next

    GetAllAssetPairs = retVal
End Function

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
    Dim sAssetPairs, sFeeInfo As String

    sAssetPairs = ActiveSheet.Range(cellInputAssetPairs).Value
    If Len(sAssetPairs) > 0 Then
        params.Add "pair", sAssetPairs
    Else
        sAssetPairs = GetAllAssetPairs()
        ActiveSheet.Range(cellInputAssetPairs).Value = sAssetPairs
        params.Add "pair", sAssetPairs
    End If
    sFeeInfo = ActiveSheet.Range(cellInputFeeInfos).Value
    If Len(sFeeInfo) > 0 Then
        params.Add "fee-info", sFeeInfo
    End If

    Set GatherInputParameters = params
End Function

Private Sub DisplaySingleRow(ByVal key As String, ByRef item As Variant, ByRef itemRel As Variant, ByVal iRow As Integer)
    ActiveSheet.Range("B" & iRow).Value = key

    ActiveSheet.Range("C" & iRow).Value = item("fee")
    ActiveSheet.Range("D" & iRow).Value = item("minfee")
    ActiveSheet.Range("E" & iRow).Value = item("maxfee")
    ActiveSheet.Range("F" & iRow).Value = item("nextfee")
    ActiveSheet.Range("G" & iRow).Value = item("nextvolume")
    ActiveSheet.Range("H" & iRow).Value = item("tiervolume")

    If Not itemRel Is Nothing Then
        ActiveSheet.Range("I" & iRow).Value = itemRel("fee")
        ActiveSheet.Range("J" & iRow).Value = itemRel("minfee")
        ActiveSheet.Range("K" & iRow).Value = itemRel("maxfee")
        ActiveSheet.Range("L" & iRow).Value = itemRel("nextfee")
        ActiveSheet.Range("M" & iRow).Value = itemRel("nextvolume")
        ActiveSheet.Range("N" & iRow).Value = itemRel("tiervolume")
    End If
End Sub

Private Sub DoGetTradeVolume()
    Dim creds As Dictionary
    Dim method As String
    Dim params As Dictionary
    Dim result As Variant, errJson As Variant
    Dim i As Integer, curOffset As Integer
    Dim sValue As String, resultList As Variant, resultListRel As Variant
    Dim arrKeys() As Variant, key As String, item As Variant, itemRel As Variant

    ' load credentials
    Set creds = LoadCredentials()
    If creds Is Nothing Then
        ActiveSheet.Range(cellOutputErrMsg).Value = "No Key & Secret and no credentials file found!"
        Exit Sub
    End If

    ' prepare params
    method = "TradeVolume"

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

    ActiveSheet.Range("C18").Value = result("currency")
    ActiveSheet.Range("C19").Value = result("volume")

    If result.Exists("fees") Then
        Set resultList = result("fees")
        If result.Exists("fees_maker") Then
            Set resultListRel = result("fees_maker")
        Else
            Set resultListRel = New Dictionary

            ' because the rest would break ...
            'Set resultList = result("fees2") ' error
            'resultList = result("fees") ' error
        End If
        arrKeys = resultList.Keys

        ' output rows
        For i = 0 To UBound(arrKeys)
            If i Mod 10 = 0 Then VBA.DoEvents

            key = arrKeys(i)
            Set item = resultList(key)
            If resultListRel.Exists(key) Then
                Set itemRel = resultListRel(key)
            Else
                Set itemRel = Nothing
            End If

            curOffset = offsetTopResults + i

            DisplaySingleRow key, item, itemRel, curOffset
        Next
    End If
End Sub

Private Sub DoClear()
    ' TODO: Q'n'D clearing ...
    ActiveSheet.Range(rangeResultsToClear).ClearContents
    ActiveSheet.Range(cellOutputCurrency).ClearContents
    ActiveSheet.Range(cellOutputVolume).ClearContents
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
        DoGetTradeVolume
        Application.ScreenUpdating = True
    ElseIf cellLinkClearBtn = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoClear
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################
