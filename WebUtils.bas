Attribute VB_Name = "WebUtils"
Option Explicit

Private msXML As XMLHTTP60
Private msXMLPost As ServerXMLHTTP60

' Inside the VBE, Go to Tools -> References, then Select Microsoft XML, v6.0
' (or whatever your latest is. This will give you access to the XML Object Library.)
' Also add "Microsoft Scripting Runtime"

' #############################################################################

' JSON Parsing from: https://stackoverflow.com/a/30494373/9360161
' Download from: https://stackoverflow.com/a/52903455/9360161


' #############################################################################

Public Sub ParseJson(ByVal strContent As String, varJson As Variant, strState As String)
    ' strContent - source JSON string
    ' varJson - created object or array to be returned as result
    ' strState - Object|Array|Error depending on processing to be returned as state
    Dim objTokens As Object
    Dim objRegEx As Object
    Dim bMatched As Boolean

    Set objTokens = CreateObject("Scripting.Dictionary")
    Set objRegEx = CreateObject("VBScript.RegExp")
    With objRegEx
        ' specification http://www.json.org/
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = """(?:\\""|[^""])*""(?=\s*(?:,|\:|\]|\}))"
        Tokenize objTokens, objRegEx, strContent, bMatched, "str"
        .Pattern = "(?:[+-])?(?:\d+\.\d*|\.\d+|\d+)e(?:[+-])?\d+(?=\s*(?:,|\]|\}))"
        Tokenize objTokens, objRegEx, strContent, bMatched, "num"
        .Pattern = "(?:[+-])?(?:\d+\.\d*|\.\d+|\d+)(?=\s*(?:,|\]|\}))"
        Tokenize objTokens, objRegEx, strContent, bMatched, "num"
        .Pattern = "\b(?:true|false|null)(?=\s*(?:,|\]|\}))"
        Tokenize objTokens, objRegEx, strContent, bMatched, "cst"
        .Pattern = "\b[A-Za-z_]\w*(?=\s*\:)" ' unspecified name without quotes
        Tokenize objTokens, objRegEx, strContent, bMatched, "nam"
        .Pattern = "\s"
        strContent = .Replace(strContent, "")
        .MultiLine = False
        Do
            bMatched = False
            .Pattern = "<\d+(?:str|nam)>\:<\d+(?:str|num|obj|arr|cst)>"
            Tokenize objTokens, objRegEx, strContent, bMatched, "prp"
            .Pattern = "\{(?:<\d+prp>(?:,<\d+prp>)*)?\}"
            Tokenize objTokens, objRegEx, strContent, bMatched, "obj"
            .Pattern = "\[(?:<\d+(?:str|num|obj|arr|cst)>(?:,<\d+(?:str|num|obj|arr|cst)>)*)?\]"
            Tokenize objTokens, objRegEx, strContent, bMatched, "arr"
        Loop While bMatched
        .Pattern = "^<\d+(?:obj|arr)>$" ' unspecified top level array
        If Not (.Test(strContent) And objTokens.Exists(strContent)) Then
            varJson = Null
            strState = "Error"
        Else
            Retrieve objTokens, objRegEx, strContent, varJson
            strState = IIf(IsObject(varJson), "Object", "Array")
        End If
    End With
End Sub

Private Sub Tokenize(objTokens, objRegEx, strContent, bMatched, strType)
    Dim strKey As String
    Dim strRes As String
    Dim lngCopyIndex As Long
    Dim objMatch As Object

    strRes = ""
    lngCopyIndex = 1
    With objRegEx
        For Each objMatch In .Execute(strContent)
            strKey = "<" & objTokens.count & strType & ">"
            bMatched = True
            With objMatch
                objTokens(strKey) = .Value
                strRes = strRes & Mid(strContent, lngCopyIndex, .FirstIndex - lngCopyIndex + 1) & strKey
                lngCopyIndex = .FirstIndex + .Length + 1
            End With
        Next
        strContent = strRes & Mid(strContent, lngCopyIndex, Len(strContent) - lngCopyIndex + 1)
    End With
End Sub

Private Sub Retrieve(objTokens, objRegEx, strTokenKey, varTransfer)
    Dim strContent As String
    Dim strType As String
    Dim objMatches As Object
    Dim objMatch As Object
    Dim strName As String
    Dim varValue As Variant
    Dim objArrayElts As Object

    strType = Left(Right(strTokenKey, 4), 3)
    strContent = objTokens(strTokenKey)
    'Debug.Print "Retrieve", strType, strContent
    With objRegEx
        .Global = True
        Select Case strType
            Case "obj"
                .Pattern = "<\d+\w{3}>"
                Set objMatches = .Execute(strContent)
                Set varTransfer = CreateObject("Scripting.Dictionary")
                For Each objMatch In objMatches
                    Retrieve objTokens, objRegEx, objMatch.Value, varTransfer
                Next
            Case "prp"
                .Pattern = "<\d+\w{3}>"
                Set objMatches = .Execute(strContent)

                Retrieve objTokens, objRegEx, objMatches(0).Value, strName
                Retrieve objTokens, objRegEx, objMatches(1).Value, varValue
                If IsObject(varValue) Then
                    Set varTransfer(strName) = varValue
                Else
                    varTransfer(strName) = varValue
                End If
            Case "arr"
                .Pattern = "<\d+\w{3}>"
                Set objMatches = .Execute(strContent)
                Set objArrayElts = CreateObject("Scripting.Dictionary")
                For Each objMatch In objMatches
                    Retrieve objTokens, objRegEx, objMatch.Value, varValue
                    If IsObject(varValue) Then
                        Set objArrayElts(objArrayElts.count) = varValue
                    Else
                        objArrayElts(objArrayElts.count) = varValue
                    End If
                    varTransfer = objArrayElts.Items
                Next
            Case "nam"
                varTransfer = strContent
            Case "str"
                varTransfer = Mid(strContent, 2, Len(strContent) - 2)
                varTransfer = Replace(varTransfer, "\""", """")
                varTransfer = Replace(varTransfer, "\\", "\")
                varTransfer = Replace(varTransfer, "\/", "/")
                varTransfer = Replace(varTransfer, "\b", Chr(8))
                varTransfer = Replace(varTransfer, "\f", Chr(12))
                varTransfer = Replace(varTransfer, "\n", vbLf)
                varTransfer = Replace(varTransfer, "\r", vbCr)
                varTransfer = Replace(varTransfer, "\t", vbTab)
                .Global = False
                .Pattern = "\\u[0-9a-fA-F]{4}"
                Do While .Test(varTransfer)
                    varTransfer = .Replace(varTransfer, ChrW(("&H" & Right(.Execute(varTransfer)(0).Value, 4)) * 1))
                Loop
            Case "num"
                varTransfer = Evaluate(strContent)
            Case "cst"
                Select Case LCase(strContent)
                    Case "true"
                        varTransfer = True
                    Case "false"
                        varTransfer = False
                    Case "null"
                        varTransfer = Null
                End Select
        End Select
    End With
End Sub

Public Function BeautifyJson(varJson As Variant) As String
    Dim strResult As String
    Dim lngIndent As Long
    BeautifyJson = ""
    lngIndent = 0
    BeautyTraverse BeautifyJson, lngIndent, varJson, vbTab, 1
End Function

Private Sub BeautyTraverse(strResult As String, lngIndent As Long, varElement As Variant, strIndent As String, lngStep As Long)
    Dim arrKeys() As Variant
    Dim lngIndex As Long
    Dim strTemp As String

    Select Case VarType(varElement)
        Case vbObject
            If varElement.count = 0 Then
                strResult = strResult & "{}"
            Else
                strResult = strResult & "{" & vbCrLf
                lngIndent = lngIndent + lngStep
                arrKeys = varElement.Keys
                For lngIndex = 0 To UBound(arrKeys)
                    strResult = strResult & String(lngIndent, strIndent) & """" & arrKeys(lngIndex) & """" & ": "
                    BeautyTraverse strResult, lngIndent, varElement(arrKeys(lngIndex)), strIndent, lngStep
                    If Not (lngIndex = UBound(arrKeys)) Then strResult = strResult & ","
                    strResult = strResult & vbCrLf
                Next
                lngIndent = lngIndent - lngStep
                strResult = strResult & String(lngIndent, strIndent) & "}"
            End If
        Case Is >= vbArray
            If UBound(varElement) = -1 Then
                strResult = strResult & "[]"
            Else
                strResult = strResult & "[" & vbCrLf
                lngIndent = lngIndent + lngStep
                For lngIndex = 0 To UBound(varElement)
                    strResult = strResult & String(lngIndent, strIndent)
                    BeautyTraverse strResult, lngIndent, varElement(lngIndex), strIndent, lngStep
                    If Not (lngIndex = UBound(varElement)) Then strResult = strResult & ","
                    strResult = strResult & vbCrLf
                Next
                lngIndent = lngIndent - lngStep
                strResult = strResult & String(lngIndent, strIndent) & "]"
            End If
        Case vbInteger, vbLong, vbSingle, vbDouble
            strResult = strResult & varElement
        Case vbNull
            strResult = strResult & "Null"
        Case vbBoolean
            strResult = strResult & IIf(varElement, "True", "False")
        Case Else
            strTemp = Replace(varElement, "\""", """")
            strTemp = Replace(strTemp, "\", "\\")
            strTemp = Replace(strTemp, "/", "\/")
            strTemp = Replace(strTemp, Chr(8), "\b")
            strTemp = Replace(strTemp, Chr(12), "\f")
            strTemp = Replace(strTemp, vbLf, "\n")
            strTemp = Replace(strTemp, vbCr, "\r")
            strTemp = Replace(strTemp, vbTab, "\t")
            strResult = strResult & """" & strTemp & """"
    End Select
End Sub

' #############################################################################

Function DoHTTPGet(url As String, Optional params As Object = Nothing, Optional headers As Object = Nothing) As String
    If msXML Is Nothing Then Set msXML = New XMLHTTP60

    Dim queryString As String
    If Not params Is Nothing Then
        queryString = "?" & encodeCollection(params)
        url = url & queryString
    End If

    With msXML
        .Open "GET", url, False

        If Not headers Is Nothing Then
            Dim i As Integer
            Dim k, v

            k = headers.Keys
            v = headers.Items

            For i = 0 To headers.count - 1
                Debug.Print i, k(i), v(i), TypeName(v(i))
                If Len(CStr(v(i))) = 0 Then
                    Debug.Print "Filter out header with empty value: " & k(i)
                    Debug.Assert Len(CStr(v(i))) > 0
                Else
                    .setRequestHeader k(i), v(i)
                End If
            Next
        End If

        .send

        DoHTTPGet = StrConv(.responseBody, vbUnicode)
    End With
End Function

Function DoHTTPPost(url As String, data As Object, headers As Object) As String
    If msXMLPost Is Nothing Then Set msXMLPost = New ServerXMLHTTP60

    Debug.Assert "Dictionary" = TypeName(headers)

    Dim postData As String
    Dim i As Integer
    Dim k, v

    postData = encodeCollection(data)

    With msXMLPost
        .Open "POST", url, False

        k = headers.Keys
        v = headers.Items

        For i = 0 To headers.count - 1
            Debug.Print i, k(i), v(i), TypeName(v(i))
            If Len(CStr(v(i))) = 0 Then
                Debug.Print "Filter out header with empty value: " & k(i)
                Debug.Assert Len(CStr(v(i))) > 0
            Else
                .setRequestHeader k(i), v(i)
            End If
        Next

        .send postData

        DoHTTPPost = StrConv(.responseBody, vbUnicode)
    End With
End Function

Public Function encodeCollection(ByVal params As Object) As String
    Dim result As String
    Dim i As Integer
    Dim k, v

    If params Is Nothing Then
        encodeCollection = ""
        Exit Function
    End If

    Debug.Assert "Dictionary" = TypeName(params)

    result = ""
    k = params.Keys
    v = params.Items

    For i = 0 To params.count - 1
        If i > 0 Then
            result = result & "&"
        End If
        result = result & k(i) & "=" & v(i)
    Next

    encodeCollection = result
End Function

' #############################################################################

