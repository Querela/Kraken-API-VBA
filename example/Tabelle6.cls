VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputAssetPair As String = "C6"
Private Const cellInputInterval As String = "C8"
Private Const cellInputSince As String = "C10"
Private Const cellLinkButton As String = "B12"
Private Const cellOutputErrMsg As String = "D12"
Private Const offsetTopResults As Integer = 16
Private Const rangeResultsToClear As String = "B16:J800"  'J735

' #############################################################################
' Utils

' #############################################################################

Private Sub DoGetOHLCData()
    Dim method As String
    Dim params As Object
    Dim result As Variant, errJson As Variant
    Dim assetpair As String, sInterval As String, sSince As String
    Dim i As Integer, curOffset As Integer, arrKeys() As Variant, key As String, data As Variant, item As Variant
    
    Set params = CreateObject("Scripting.Dictionary")
    
    ' prepare params
    method = "OHLC"
    
    ' Check if assetpair given
    assetpair = ActiveSheet.Range(cellInputAssetPair).Value
    If VBA.Len(assetpair) > 2 Then
        params.Add "pair", assetpair
    End If
    sInterval = ActiveSheet.Range(cellInputInterval).Value
    If VBA.Len(sInterval) > 0 Then
        params.Add "interval", sInterval
    End If
    sSince = ActiveSheet.Range(cellInputSince).Value
    If VBA.Len(sSince) > 0 Then
        params.Add "since", sSince
    End If
    
    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPublic(method, params)
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
    arrKeys = result.Keys

    ' update since for new query
    ActiveSheet.Range(cellInputSince).Value = result("last")
    
    ' append (?) data
    key = arrKeys(0)
    data = result(key)
    
    For i = 0 To UBound(data)
        item = data(i)
        curOffset = offsetTopResults + i

        ActiveSheet.Range("B" & curOffset).Value = key
        ActiveSheet.Range("C" & curOffset).Value = item(0)
        ActiveSheet.Range("D" & curOffset).Value = item(1)
        ActiveSheet.Range("E" & curOffset).Value = item(2)
        ActiveSheet.Range("F" & curOffset).Value = item(3)
        ActiveSheet.Range("G" & curOffset).Value = item(4)
        ActiveSheet.Range("H" & curOffset).Value = item(5)
        ActiveSheet.Range("I" & curOffset).Value = item(6)
        ActiveSheet.Range("J" & curOffset).Value = item(7)
    Next
End Sub


' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String
    
    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get OHLC Data ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetOHLCData
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################



