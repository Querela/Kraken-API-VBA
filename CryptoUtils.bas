Attribute VB_Name = "CryptoUtils"
Option Explicit

Private msXML As XMLHTTP60
Private msXMLPost As ServerXMLHTTP60

' Inside the VBE, Go to Tools -> References, then Select Microsoft XML, v6.0
' (or whatever your latest is. This will give you access to the XML Object Library.)
' Also add "Microsoft Scripting Runtime"

' #############################################################################

' Crypto from: https://www.excelhowto.com/macros/excel-vba-base64-hmac-sha256-and-sha1-encryption/
' And: https://stackoverflow.com/questions/8246340/does-vba-have-a-hash-hmac
' Hashing from: https://microsoft-programmierer.de/Details_Mobile?d=2978&a=8&f=165&l=0&v=m&t=Excel-:-Werte-kodieren-mit-HASH-Funktionen-SHA256


' #############################################################################

Public Function HMACSHA512(ByRef bData() As Byte, ByVal sSharedSecretKey As String) As Byte()
    Dim enc As Object
    Dim SharedSecretKey() As Byte

    Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")

    SharedSecretKey = ToBytes(sSharedSecretKey)
    enc.key = SharedSecretKey

    HMACSHA512 = enc.ComputeHash_2(bData)

    Set enc = Nothing
End Function

Public Function HMACSHA512_2(ByRef bData() As Byte, ByRef bSharedSecretKey() As Byte) As Byte()
    Dim enc As Object

    Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")
    enc.key = bSharedSecretKey

    HMACSHA512_2 = enc.ComputeHash_2(bData)

    Set enc = Nothing
End Function

Public Function SHA256(bInput() As Byte) As Byte()
    Dim Encoder_SHA256 As Object
    Set Encoder_SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    SHA256 = Encoder_SHA256.ComputeHash_2(bInput)

    Set Encoder_SHA256 = Nothing
End Function

' #############################################################################

Public Function ToBytes(ByVal data As Variant) As Byte()
    Dim Encoder As Object
    Set Encoder = CreateObject("System.Text.UTF8Encoding")

    ToBytes = Encoder.GetBytes_4(data)

    Set Encoder = Nothing
End Function

Public Function FromBytes(ByRef bData() As Byte) As String
    Dim Decoder As Object
    Set Decoder = CreateObject("System.Text.UTF8Encoding")

    FromBytes = Decoder.GetString(bData)

    Set Decoder = Nothing
End Function

Public Function ByteConcat(ByRef bArr1() As Byte, ByRef bArr2() As Byte) As Byte()
    Dim i As Long
    Dim bArrOut() As Byte

    bArrOut = bArr1
    ReDim Preserve bArrOut(UBound(bArr1) + UBound(bArr2) + 1)

    For i = 0 To UBound(bArr2)
        bArrOut(i + UBound(bArr1) + 1) = bArr2(i)
    Next

    ByteConcat = bArrOut
End Function

Public Function BytesHexString(ByRef bData() As Byte) As String
    ' From: http://www.vbforums.com/showthread.php?559398-Byte-array-to-hex-string
    Dim i As Long, sHex As String

    For i = LBound(bData) To UBound(bData)
        sHex = sHex & Hex$(bData(i)) & " "
    Next

    'Remove last space at end.
    BytesHexString = Left$(sHex, Len(sHex) - 1)
End Function

' #############################################################################

Public Function SHA256String(sInput As String, Optional bB64 As Boolean = 0) As String
    Dim bytes() As Byte

    bytes = ToBytes(sInput)
    bytes = SHA256(bytes)

    If bB64 Then
        SHA256String = ConvToBase64String(bytes)
    Else
        SHA256String = ConvToHexString(bytes)
    End If
End Function

' #############################################################################

Public Function EncodeBase64(ByRef arrData() As Byte, Optional ByVal noLinebreaks As Boolean = True) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMNode
    Dim sDecoded As String

    Set objXML = New MSXML2.DOMDocument60

    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData

    sDecoded = objNode.Text
    If noLinebreaks Then
        sDecoded = Replace(sDecoded, vbLf, "")
    End If
    EncodeBase64 = sDecoded

    Set objNode = Nothing
    Set objXML = Nothing
End Function

Public Function DecodeBase64(ByRef sEncoded As String) As Byte()
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMNode

    Set objXML = New MSXML2.DOMDocument60

    ' base64 to byte array
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = sEncoded

    DecodeBase64 = objNode.nodeTypedValue

    Set objNode = Nothing
    Set objXML = Nothing
End Function

Public Function ConvToBase64String(vIn As Variant) As String
    Dim objXML As MSXML2.DOMDocument60
    Set objXML = New MSXML2.DOMDocument60

    With objXML
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
    End With

    ConvToBase64String = Replace(objXML.DocumentElement.Text, vbLf, "")

    Set objXML = Nothing
End Function

Public Function ConvToHexString(vIn As Variant) As String
    'Dim oD As Object
    'Set oD = CreateObject("MSXML2.DOMDocument")
    Dim objXML As MSXML2.DOMDocument60
    Set objXML = New MSXML2.DOMDocument60

    With objXML
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
    End With

    ConvToHexString = Replace(objXML.DocumentElement.Text, vbLf, "")

    Set objXML = Nothing
End Function

' #############################################################################

