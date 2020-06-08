Attribute VB_Name = "Test"
Option Explicit

Public Sub Test_General()
    Debug.Print "Test Encryption general"
    
    Dim data As String
    Dim secret As String
    Dim result() As Byte

    data = "abc"
    secret = "123"

    result = CryptoUtils.SHA256String(data)
    Debug.Print result
    result = CryptoUtils.SHA256String(data, True)
    Debug.Print result

    ' bytearray(hmac.new(b"123", b"abc", hashlib.sha512).digest()).hex()
    result = CryptoUtils.ConvToHexString(CryptoUtils.HMACSHA512(CryptoUtils.ToBytes(data), secret))
    Debug.Print result
End Sub

Public Sub Test_KrakenSign()
    Debug.Print "Test Kraken Signing"
    
    Dim data As String
    Dim secret As String
    Dim path, method As String
    Dim result() As Byte
    Dim result2 As String

    'Dim params As Collection
    'Set params = New Collection
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    
    Debug.Print "nonce?", API.nonce()
    params.Add "nonce", "1591386544485"
    Debug.Print "params", WebUtils.encodeCollection(params)

    data = "abc"
    secret = "123"
    method = ""
    path = "/0/private/" + method
    
    result2 = API.KrakenSign(secret, params, path)
    Debug.Print result2
    ' message=b'/0/private/\xa2\xaa\t\x93;\xc2i\xa7E\x89\xb5F\xf8\xe0\xbc\xcb]>]*\x1e\xa0A\xed"\x8f@f\xcfJI\xec'
    ' sigdigest=b'1D2VOXmMaXoywj/KO7WvD4Q9KckvkjVnDaVvDjKF4udJTP/6VW+veDTJUPGiux7Ulyu5ix8SKVRhD4kk77sDmw=='
    
    ' headers = {
    '     'API-Key': self.key,
    '     'API-Sign': self._sign(data, urlpath)
    ' }
End Sub

Public Sub Test_Post()
    ' DoHTTPPost
    
    Dim headers As Object, data As Object
    Dim uri As String, urlpath As String, signature As String
    Dim secret As String, key As String
    Dim result As String
    
    Set headers = CreateObject("Scripting.Dictionary")
    Set data = CreateObject("Scripting.Dictionary")
    
    key = "abc"
    secret = "123"
    
    'Debug.Print "nonce?", KrakenUtils.nonce()
    data.Add "nonce", "1591386544485"
    
    urlpath = "/0/private/"
    signature = API.KrakenSign(secret, data, urlpath)
    
    headers.Add "API-Key", key
    headers.Add "API-Sign", signature
    
    ' Adjust for POST testing
    urlpath = "/post"
    uri = "https://httpbin.org" & urlpath
    
    result = WebUtils.DoHTTPPost(uri, data, headers)
    Debug.Print "result", result
    
End Sub

Public Sub Test_TypeCheck()
    Dim c As Object
    Debug.Print TypeName(c)
    
    If c Is Nothing Then
        Debug.Print "c should be Nothing"
    End If
    If "Dictionary" = TypeName(c) Then
        Debug.Print "c shouldn't be be a Dictionary"
    End If
    
    Set c = CreateObject("Scripting.Dictionary")
    Debug.Print TypeName(c)
    
    If "Dictionary" = TypeName(c) Then
        Debug.Print "c should now be a Dictionary"
    End If
End Sub

Public Sub Test_KrakenPrivate1()
    Dim key, secret As String
    Dim method As String
    Dim result As Variant

    ' Note key should not be empty, as empty headers are not allowed?!?
    key = "key"
    secret = "secret"
    method = "Balance"
    
    ' invalid credentials
    Set result = API.KrakenQueryPrivate(key, secret, method)
    Debug.Print method, WebUtils.BeautifyJson(result)
    
    ' Invalid method
    method = "Balance2"
    Set result = API.KrakenQueryPrivate(key, secret, method)
    Debug.Print method, WebUtils.BeautifyJson(result)
    
    ' Invalid uri
    ' only if we forget to concat the kraken base api with the urlpath ...
End Sub

Public Sub Test_KrakenPublic1()
    Dim method As String
    Dim result As Variant
    Dim data As Object
    
    method = "Ticker"
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "pair", "XXBTZEUR"
    
    Set result = API.KrakenQueryPublic(method, data)
    Debug.Print method, WebUtils.BeautifyJson(result)
End Sub

Public Sub Test_IO2()
    Dim creds As Object
    Set creds = FileUtils.LoadKrakenCredentials
    Debug.Print "Creds", WebUtils.BeautifyJson(creds)
    Debug.Print "The secret is """ & creds("secret") & """."
End Sub

Public Sub Test_decodeBase64()
    Dim creds As Object

    Set creds = FileUtils.LoadKrakenCredentials
    Debug.Print "Creds", WebUtils.BeautifyJson(creds)
    Debug.Print "The secret is """ & creds("secret") & """."

    Dim bSecret() As Byte, sSecret As String, sSecret2 As String

    sSecret = creds("secret")

    bSecret = CryptoUtils.DecodeBase64(sSecret)
    sSecret2 = CryptoUtils.EncodeBase64(bSecret)
    Debug.Assert sSecret = sSecret2
    Debug.Print "secret (orig)", sSecret
    Debug.Print "secret (dupl)", sSecret2

    Debug.Print CryptoUtils.EncodeBase64(bSecret)
    Debug.Print CryptoUtils.EncodeBase64(bSecret, noLinebreaks:=False)
    Debug.Print CryptoUtils.ConvToHexString(bSecret)
    Debug.Print CryptoUtils.ConvToBase64String(bSecret)
End Sub

Public Sub Test_KrakenPrivate2()
    Dim creds As Object
    Dim key, secret As String
    Dim method As String
    Dim result As Variant

    ' TODO: how to handle missing file?
    Set creds = FileUtils.LoadKrakenCredentials
    Debug.Print "Creds", WebUtils.BeautifyJson(creds)

    ' Note key should not be empty, as empty headers are not allowed?!?
    key = creds("key")
    secret = creds("secret")
    method = "Balance"

    Set result = API.KrakenQueryPrivate(key, secret, method)
    Debug.Print method, WebUtils.BeautifyJson(result)
End Sub

Public Sub Test_nonces()
    Dim nonce As String
    nonce = CStr((Now() - 25569) * 86400)
    Debug.Print "nonce", nonce
    Debug.Print "Timer", Timer
    Debug.Print "timestamp?", DateDiff("s", "01/01/1970", Date)
    Debug.Print "better1?", (DateDiff("s", "01/01/1970", Date) + Timer) * 1000
    Debug.Print "better2?", Split((DateDiff("s", "01/01/1970", Date) + Timer) * 1000, ".")(0)
    Debug.Print "better2?", Split((DateDiff("s", "01/01/1970", Date) + Timer) * 1000, ",")(0)
    Debug.Print "better4?", CLngLng((DateDiff("s", "01/01/1970", Date) + Timer) * 1000)
End Sub
