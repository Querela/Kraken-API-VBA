VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellLinkNonce As String = "B6"
Private Const cellOutputNonce As String = "C6"

Private Const cellLinkB64Encode As String = "B14"
Private Const cellInputB64Encode As String = "C14"
Private Const cellOutputB64Encode As String = "D14"
Private Const cellLinkB64Decode As String = "B15"
Private Const cellInputB64Decode As String = "C15"
Private Const cellOutputB64Decode As String = "D15"

Private Const cellLinkSHA256 As String = "B17"
Private Const cellInputSHA256 As String = "C17"
Private Const cellOutputSHA256 As String = "D17"

Private Const cellLinkB64ToHex As String = "B19"
Private Const cellInputB64ToHex As String = "C19"
Private Const cellOutputB64ToHex As String = "D19"

Private Const cellLinkUpdateAll As String = "B4"
Private Const cellLinkClearBtn As String = "A1"

' #############################################################################

' #############################################################################

Private Sub DoUpdateNonce()
    Dim nonce As String
    nonce = API.nonce_3()
    With ActiveSheet.Range(cellOutputNonce)
        .Value = nonce
        .NumberFormat = xlNumberFormatTypeDefault
    End With
End Sub

Private Sub DoEncodeBase64()
    Dim sIn, sOut As String
    Dim bIn() As Byte

    ActiveSheet.Range(cellOutputB64Encode).ClearContents

    sIn = ActiveSheet.Range(cellInputB64Encode).Value
    bIn = CryptoUtils.ToBytes(sIn)
    sOut = CryptoUtils.EncodeBase64(bIn, True)
    ActiveSheet.Range(cellOutputB64Encode).Value = sOut
End Sub

Private Sub DoDecodeBase64()
    Dim sIn As String
    Dim sOut As String
    Dim bOut() As Byte

    ActiveSheet.Range(cellOutputB64Decode).ClearContents

    sIn = ActiveSheet.Range(cellInputB64Decode).Value
    bOut = CryptoUtils.DecodeBase64(sIn)
    'sOut = CStr(bOut)
    sOut = CryptoUtils.FromBytes(bOut)
    ActiveSheet.Range(cellOutputB64Decode).Value = sOut
End Sub

Private Sub DoSHA256()
    Dim sIn, sOut As String
    Dim bIn() As Byte
    Dim bOut() As Byte

    ActiveSheet.Range(cellOutputSHA256).ClearContents

    sIn = ActiveSheet.Range(cellInputSHA256).Value
    bIn = CryptoUtils.ToBytes(sIn)
    bOut = CryptoUtils.SHA256(bIn)
    ' sOut = CryptoUtils.FromBytes(bOut)
    sOut = CryptoUtils.EncodeBase64(bOut)
    ActiveSheet.Range(cellOutputSHA256).Value = sOut
End Sub

Private Sub DoB64ToHex()
    Dim sIn As String
    Dim sOut As String
    Dim bIn() As Byte

    ActiveSheet.Range(cellOutputB64ToHex).ClearContents

    sIn = ActiveSheet.Range(cellInputB64ToHex).Value
    bIn = CryptoUtils.DecodeBase64(sIn)
    sOut = CryptoUtils.BytesHexString(bIn)
    ActiveSheet.Range(cellOutputB64ToHex).Value = sOut
End Sub

Private Sub DoClear()
    Debug.Assert False
End Sub

' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String
    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkNonce = cellName Then
        Debug.Print "Get Nonce ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoUpdateNonce
        Application.ScreenUpdating = True
    ElseIf cellLinkB64Encode = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoEncodeBase64
        Application.ScreenUpdating = True
    ElseIf cellLinkB64Decode = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoDecodeBase64
        Application.ScreenUpdating = True
    ElseIf cellLinkSHA256 = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoSHA256
        Application.ScreenUpdating = True
    ElseIf cellLinkB64ToHex = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoB64ToHex
        Application.ScreenUpdating = True
    ElseIf cellLinkUpdateAll = cellName Then
        Application.ScreenUpdating = False
        'Application.ActiveWorkbook.ForceFullCalculation = True
        ActiveSheet.Calculate
        'Application.ActiveWorkbook.ForceFullCalculation = False
        ' won't be disabled ...
        Application.ScreenUpdating = True
    ElseIf cellLinkClearBtn = cellName Then
        Application.ScreenUpdating = False
        On Error Resume Next
        DoClear
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################

