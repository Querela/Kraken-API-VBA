VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputAssets As String = "C6"
Private Const cellLinkButton As String = "B8"
Private Const cellOutputErrMsg As String = "D8"
Private Const offsetTopResults As Integer = 12
Private Const rangeResultsToClear As String = "B12:F54"

' #############################################################################
' Utils

' #############################################################################

Private Sub DoGetAssetInfos()
    Dim method As String
    Dim params As Object
    Dim result As Variant, item As Variant, errJson As Variant
    Dim assets As String
    Dim i As Integer, curOffset As Integer, arrKeys() As Variant, key As String

    Set params = CreateObject("Scripting.Dictionary")

    ' prepare params
    method = "Assets"

    ' Check if assets given
    assets = ActiveSheet.Range(cellInputAssets).Value
    If VBA.Len(assets) > 2 Then
        params.Add "asset", assets
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
    For i = 0 To UBound(arrKeys)
        If i Mod 25 = 0 Then VBA.DoEvents

        key = arrKeys(i)
        Set item = result(key)
        curOffset = offsetTopResults + i

        ActiveSheet.Range("B" & curOffset).Value = key
        ActiveSheet.Range("C" & curOffset).Value = item("altname")
        ActiveSheet.Range("D" & curOffset).Value = item("aclass")
        ActiveSheet.Range("E" & curOffset).Value = item("decimals")
        ActiveSheet.Range("F" & curOffset).Value = item("display_decimals")
    Next
End Sub


' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String

    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get Asset Info ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetAssetInfos
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################

