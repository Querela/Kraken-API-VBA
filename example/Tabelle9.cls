VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputAssetPair As String = "C6"
Private Const cellInputSince As String = "C8"
Private Const cellLinkButton As String = "B10"
Private Const cellOutputErrMsg As String = "D10"
Private Const offsetTopResults As Integer = 14
Private Const rangeResultsToClear As String = "B14:E214"

' #############################################################################
' Utils

' #############################################################################

Private Sub DoGetRecentSpread()
    Dim method As String
    Dim params As Object
    Dim result As Variant, errJson As Variant
    Dim assetpair As String, sSince As String
    Dim i As Integer, curOffset As Integer
    Dim arrKeys() As Variant, key As String
    Dim data As Variant, item As Variant

    Set params = CreateObject("Scripting.Dictionary")
    
    ' prepare params
    method = "Spread"

    ' Check if assetpair given
    assetpair = ActiveSheet.Range(cellInputAssetPair).Value
    If VBA.Len(assetpair) > 2 Then
        params.Add "pair", assetpair
    End If
    sSince = ActiveSheet.Range(cellInputSince).Value
    If VBA.Len(sSince) > 2 Then
        params.Add "since", sSince
    End If

    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPublic(method, params)
    ' NOTE: DISABLED because of too long output?
    'Debug.Print method, WebUtils.BeautifyJson(result)

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
        If i Mod 25 = 0 Then VBA.DoEvents

        item = data(i)
        curOffset = offsetTopResults + i

        ActiveSheet.Range("B" & curOffset).Value = key
        ActiveSheet.Range("C" & curOffset).Value = item(0)
        ActiveSheet.Range("D" & curOffset).Value = item(1)
        ActiveSheet.Range("E" & curOffset).Value = item(2)
    Next
End Sub


' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String

    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get Recent Spread ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetRecentSpread
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################





