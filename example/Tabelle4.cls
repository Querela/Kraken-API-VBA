VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputAssetPairs As String = "C6"
Private Const cellInputInfos As String = "C8"
Private Const cellLinkButton As String = "B10"
Private Const cellOutputErrMsg As String = "D10"
Private Const offsetTopResults As Integer = 14
Private Const rangeResultsToClear As String = "B14:S200"

' #############################################################################
' Utils

' #############################################################################

Private Sub DoGetTradableAssetPairs()
    Dim method As String
    Dim params As Object
    Dim result As Variant, item As Variant, errJson As Variant
    Dim assetpairs As String, infoType As String
    Dim i As Integer, curOffset As Integer, arrKeys() As Variant, key As String
    
    Set params = CreateObject("Scripting.Dictionary")
    
    ' prepare params
    method = "AssetPairs"
    
    ' Check if assetpairs given
    assetpairs = ActiveSheet.Range(cellInputAssetPairs).Value
    If VBA.Len(assetpairs) > 2 Then
        params.Add "pair", assetpairs
    End If
    infoType = ActiveSheet.Range(cellInputInfos).Value
    If VBA.Len(infoType) > 2 Then
        params.Add "info", infoType
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
    curOffset = offsetTopResults
    For i = 0 To UBound(arrKeys)
        key = arrKeys(i)
        Set item = result(key)

        ActiveSheet.Range("B" & curOffset).Value = key
        ActiveSheet.Range("C" & curOffset).Value = item("altname")
        ActiveSheet.Range("D" & curOffset).Value = item("wsname")
        ActiveSheet.Range("E" & curOffset).Value = item("aclass_base")
        ActiveSheet.Range("F" & curOffset).Value = item("base")
        ActiveSheet.Range("G" & curOffset).Value = item("aclass_quote")
        ActiveSheet.Range("H" & curOffset).Value = item("quote")
        ActiveSheet.Range("I" & curOffset).Value = item("lot")
        ActiveSheet.Range("J" & curOffset).Value = item("pair_decimals")
        ActiveSheet.Range("K" & curOffset).Value = item("lot_decimals")
        ActiveSheet.Range("L" & curOffset).Value = item("lot_multiplier")
        ' ActiveSheet.Range("M" & curOffset).Value = item("leverage_buy")
        ' ActiveSheet.Range("N" & curOffset).Value = item("leverage_sell")
        ' ActiveSheet.Range("O" & curOffset).Value = item("fees")
        ' ActiveSheet.Range("P" & curOffset).Value = item("fees_maker")
        ActiveSheet.Range("Q" & curOffset).Value = item("fee_volume_currency")
        ActiveSheet.Range("R" & curOffset).Value = item("margin_call")
        ActiveSheet.Range("S" & curOffset).Value = item("margin_stop")
        
        curOffset = curOffset + 1
    Next
End Sub


' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String
    
    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get Tradable AssetPairs ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetTradableAssetPairs
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################


