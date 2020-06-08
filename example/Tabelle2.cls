VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const colnameAttName As String = "C"
Private Const colnameAttValue As String = "D"

Private Const offsetTopServerTime As Integer = 5

' #############################################################################

Private Sub DoGetServerTime()
    Dim method As String, url As String
    Dim params As Object
    Dim result As Variant
    Dim price As Double

    ' prepare params
    method = "Time"

    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPublic(method)
    Debug.Print method, WebUtils.BeautifyJson(result)

    ActiveSheet.Range(colnameAttName & offsetTopServerTime).Value = "rfc1123"
    ActiveSheet.Range(colnameAttValue & offsetTopServerTime).Value = result("result")("rfc1123")
    ActiveSheet.Range(colnameAttName & (offsetTopServerTime + 1)).Value = "unixtime"
    ActiveSheet.Range(colnameAttValue & (offsetTopServerTime + 1)).Value = result("result")("unixtime")
End Sub


' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    If "Get server time" = Target.name Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetServerTime
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################
