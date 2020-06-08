Attribute VB_Name = "API"
Option Explicit

' Inside the VBE, Go to Tools -> References, then Select Microsoft XML, v6.0
' (or whatever your latest is. This will give you access to the XML Object Library.)
' Also add "Microsoft Scripting Runtime"

' #############################################################################

Private Const krakenVersion As Integer = 0
Private Const krakenAPIBaseURI As String = "https://api.kraken.com"
Private Const myUserAgent As String = "VBA-KrakenEx/0.2 (https://github.com/Querela/Kraken-API-VBA)"

' #############################################################################

Public Function RetrieveCurrentPrice(ByVal name As String) As Double
    Dim url As String
    Dim strJsonString As String
    Dim varJson As Variant
    Dim strState As String
    Dim price As Double
    
    Debug.Print "Retrieve current price for:", name

    ' build URL
    url = krakenAPIBaseURI & "/" & krakenVersion & "/public/Ticker?pair=" & name
    ' download json
    strJsonString = WebUtils.DoHTTPGet(url)
    ' parse json to object
    WebUtils.ParseJson strJsonString, varJson, strState
    ' get current price
    price = Evaluate(varJson("result")(name)("c")(0))
    
    RetrieveCurrentPrice = price
    
    Debug.Print "Current price of " & name & ": " & price
End Function

Public Function RetrieveCurrentPrice2(ByVal name As String) As Double
    Dim method As String, url As String
    Dim params As Object
    Dim result As Variant
    Dim price As Double
    
    Debug.Print "Retrieve current price for:", name

    ' prepare params
    method = "Ticker"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "pair", name
    
    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPublic(method, params)
    Debug.Print method, WebUtils.BeautifyJson(result)

    ' get current price
    price = Evaluate(result("result")(name)("c")(0))
    
    RetrieveCurrentPrice2 = price
    
    Debug.Print "Current price of " & name & ": " & price
End Function

' #############################################################################

Public Function nonce() As String
    ' time in seconds
    nonce = CStr((Now() - 25569) * 86400 * 1000)
End Function

Public Function nonce_2() As String
    ' millisecond time
    nonce_2 = CStr((Now() - 25569) * 86400 * 1000)
End Function

Public Function nonce_3() As String
    ' millisecond time
    nonce_3 = CStr(CLngLng((DateDiff("s", "01/01/1970", Date) + Timer) * 1000))
End Function

' krakenex translation?

Public Function KrakenSign(ByVal sSharedSecretKey As String, ByVal data As Object, ByVal urlpath As String) As String
    '''
    'postdata = urllib.Parse.urlencode(data)
    '# Unicode-objects must be encoded before hashing
    'encoded = (str(data['nonce']) + postdata).encode()
    'Message = urlpath.encode() + hashlib.SHA256(encoded).digest()
    'signature = hmac.new(base64.b64decode(self.secret),
    '                     message, hashlib.sha512)
    'sigdigest = base64.b64encode(Signature.digest())
    'return sigdigest.decode()
    '''

    Dim postData As String
    Dim encoded() As Byte
    Dim message() As Byte
    Dim bSharedSecretKey() As Byte
    Dim sigdigest() As Byte
    Dim signature As String
    Dim b1() As Byte
    Dim b2() As Byte

    postData = WebUtils.encodeCollection(data)
    'Debug.Print "postdata", postdata
    'Debug.Print "nonce", data("nonce")
    'Debug.Print "tempStr", data("nonce") & postdata
    encoded = CryptoUtils.ToBytes(data("nonce") & postData)
    'Debug.Print "encoded", CryptoUtils.ConvToHexString(encoded)

    b1 = CryptoUtils.ToBytes(urlpath)
    b2 = CryptoUtils.SHA256(encoded)
    'Debug.Print "urlpath", urlpath
    'Debug.Print "bytes(urlpath)", CryptoUtils.ConvToHexString(b1)
    'Debug.Print "sha256(encoded)", CryptoUtils.ConvToHexString(b2)
    message = CryptoUtils.ByteConcat(b1, b2)
    Debug.Print "message", CryptoUtils.ConvToHexString(message)

    bSharedSecretKey = CryptoUtils.DecodeBase64(sSharedSecretKey)

    'sigdigest = CryptoUtils.HMACSHA512(message, sSharedSecretKey)
    sigdigest = CryptoUtils.HMACSHA512_2(message, bSharedSecretKey)
    'Debug.Print "sigdigest", CryptoUtils.ConvToHexString(sigdigest)
    signature = CryptoUtils.ConvToBase64String(sigdigest)
    Debug.Print "signature", signature

    KrakenSign = signature
End Function

' #############################################################################

'Private Function KrakenQueryBase(urlpath, data, headers) As String
'End Function

Public Function KrakenQueryPublic(ByVal sMethod As String, Optional ByVal params As Object = Nothing) As Variant
    Dim uri As String, urlpath As String, queryString As String
    Dim strJsonString As String
    Dim varJson As Variant
    Dim strState As String
    Dim i As Integer
    Dim errJson As Variant
    
    ' build URI Path
    urlpath = "/" & krakenVersion & "/public/" & sMethod
    
    ' build query string
    'If params Is Nothing Then
    '    queryString = ""
    'Else
    '    queryString = "?" & WebUtils.encodeCollection(params)
    'End If
    
    ' build URI
    uri = krakenAPIBaseURI & urlpath '& queryString

    ' download json
    strJsonString = WebUtils.DoHTTPGet(uri, params)
    ' parse json to object
    WebUtils.ParseJson strJsonString, varJson, strState
    Debug.Assert "Object" = strState

    ' check error?
    errJson = varJson("error")
    If Not IsEmpty(errJson) Then
        Debug.Assert -1 = UBound(errJson)
        If UBound(errJson) >= 0 Then
            Debug.Print "Errors for """ & urlpath & """:"
            For i = 0 To UBound(errJson)
                Debug.Print "  " & errJson(i)
            Next
        End If
    End If
    
    Set KrakenQueryPublic = varJson
End Function

Public Function KrakenQueryPrivate(ByVal sKey As String, ByVal sSecret As String, ByVal sMethod As String, Optional ByVal data As Object = Nothing) As Variant
    Dim uri As String, urlpath As String
    Dim signature As String
    Dim headers As Object
    Dim result As String
    Dim i As Integer
    
    Set headers = CreateObject("Scripting.Dictionary")
    
    If data Is Nothing Then
        Set data = CreateObject("Scripting.Dictionary")
    End If
    Debug.Assert "Dictionary" = TypeName(data)
    data.Add "nonce", nonce_3()
    
    urlpath = "/" & krakenVersion & "/private/" & sMethod
    
    signature = KrakenSign(sSecret, data, urlpath)
    
    headers.Add "User-Agent", myUserAgent
    headers.Add "API-Key", sKey
    headers.Add "API-Sign", signature
    
    uri = "https://api.kraken.com" & urlpath
    
    'uri = "https://httpbin.org/post" ' DEBUG/TESTING
    result = WebUtils.DoHTTPPost(uri, data, headers)
    KrakenQueryPrivate = result
    
    ' parse JSON
    Dim strJsonString As String
    Dim varJson As Variant
    Dim strState As String
    Dim errJson As Variant
    
    strJsonString = result
    ' parse json to object
    WebUtils.ParseJson strJsonString, varJson, strState
    Debug.Assert "Object" = strState

    ' get result item if not error?
    errJson = varJson("error")
    If Not IsEmpty(errJson) Then
        Debug.Assert -1 = UBound(errJson)
        If UBound(errJson) >= 0 Then
            Debug.Print "Errors for """ & urlpath & """:"
            For i = 0 To UBound(errJson)
                Debug.Print "  " & errJson(i)
            Next
        End If
    End If

    Set KrakenQueryPrivate = varJson
End Function

' #############################################################################
' #############################################################################
' #############################################################################

' #############################################################################
' #############################################################################
' #############################################################################

' ''''''''''''''''''''''''''''''
' Links
' ''''''''''''''''''''''''''''''

' https://docs.microsoft.com/de-de/office/vba/library-reference/concepts/getting-started-with-vba-in-office
' https://support.kraken.com/hc/en-us/articles/360000919986-Public-endpoint-examples-you-can-try-them-directly-in-a-web-browser-
' https://www.kraken.com/features/api#public-market-data
' https://api.kraken.com/0/public/Ticker?pair=xbteur
' https://stackoverflow.com/questions/817602/gethttp-with-excel-vba
' http://excelerator.solutions/2017/08/28/excel-http-get-request/
' https://stackoverflow.com/questions/19360440/how-to-parse-json-with-vba-without-external-libraries
' https://stackoverflow.com/questions/6627652/parsing-json-in-excel-vba
' https://stackoverflow.com/questions/16817545/handle-json-object-in-xmlhttp-response-in-excel-vba-code/16851758#16851758
' https://stackoverflow.com/questions/3872339/what-is-the-difference-between-dim-and-set-in-vba
' https://excelmacromastery.com/excel-vba-range-cells/
' https://bettersolutions.com/excel/cells-ranges/vba-code.htm
' https://stackoverflow.com/questions/35318253/get-the-value-of-a-cell-in-range-store-it-a-variable-then-next-cell-next-row
' https://docs.microsoft.com/de-de/office/vba/api/excel.worksheetfunction.counta
' https://docs.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/looping-through-a-range-of-cells
' https://github.com/OfficeDev/VBA-content/blob/master/VBA/Office-Shared-VBA/articles/getting-started-with-vba-in-office.md
' https://github.com/OfficeDev/VBA-content/blob/master/VBA/Language-Reference-VBA/readme.md
' https://www.spreadsheetsmadeeasy.com/getting-and-setting-cell-values-vba/

' https://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
' https://excelmacromastery.com/vba-dictionary/
' http://excelerator.solutions/2017/08/28/excel-http-get-request/
' https://www.tek-tips.com/viewthread.cfm?qid=1469674
' https://stackoverflow.com/questions/1703505/excel-date-to-unix-timestamp
' https://www.di-mgt.com.au/binary-and-byte-operations-in-VB6.html
' https://stackoverflow.com/questions/46394144/vba-concatenate-byte-array
' https://docs.microsoft.com/de-de/office/vba/language/reference/user-interface-help/dictionary-object
' https://stackoverflow.com/questions/11463799/what-is-the-best-vba-data-typekey-value-to-save-data-same-as-php-array
' https://stackoverflow.com/questions/53558110/xml-parse-vba-excel-function-trip-msxml2-domdocument/53559474
' https://stackoverflow.com/questions/8246340/does-vba-have-a-hash-hmac
' https://stackoverflow.com/questions/50449004/convert-an-array-of-bytes-into-a-string
' https://stackoverflow.com/questions/33941363/determining-the-full-type-of-a-variable
' https://docs.microsoft.com/de-de/office/vba/api/excel.application.screenupdating
'
' https://stackoverflow.com/questions/7004754/how-to-programmatically-code-an-undo-function-in-excel-vba
' https://www.jkp-ads.com/Articles/UndoWithVBA00.asp
' https://github.com/VBA-tools/VBA-Web/blob/master/src/WebHelpers.bas#L1438
' https://www.extendoffice.com/documents/excel/2473-excel-timestamp-to-date.html
' https://www.contextures.com/xlDataVal01.html#create
' https://www.mrexcel.com/board/threads/how-to-get-unix-timestamp-in-milliseconds-vba.973463/
'

' #############################################################################

