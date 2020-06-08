VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

Private Const cellInputAssetPairs As String = "C6"
Private Const cellLinkButton As String = "B8"
Private Const cellOutputErrMsg As String = "D8"
Private Const offsetTopResults As Integer = 12
Private Const rangeResultsToClear As String = "B12:U200"

' unused
Private Const allAssetPairs As String = "" _
    & "ADAETH,ADAEUR,ADAUSD,ADAXBT,ALGOETH,ALGOEUR,ALGOUSD,ALGOXBT,ATOMETH," _
    & "ATOMEUR,ATOMUSD,ATOMXBT,BATETH,BATEUR,BATUSD,BATXBT,BCHETH,BCHEUR," _
    & "BCHGBP,BCHUSD,BCHUSDT,BCHXBT,DAIEUR,DAIUSD,DAIUSDT,DASHEUR,DASHUSD," _
    & "DASHXBT,EOSETH,EOSEUR,EOSUSD,EOSXBT,ETHCHF,ETHDAI,ETHUSDC,ETHUSDT," _
    & "EURCAD,EURCHF,EURGBP,EURJPY,GNOETH,GNOEUR,GNOUSD,GNOXBT,ICXETH,ICXEUR," _
    & "ICXUSD,ICXXBT,LINKETH,LINKEUR,LINKUSD,LINKXBT,LSKETH,LSKEUR,LSKUSD," _
    & "LSKXBT,LTCETH,LTCGBP,LTCUSDT,NANOETH,NANOEUR,NANOUSD,NANOXBT,OMGETH," _
    & "OMGEUR,OMGUSD,OMGXBT,OXTETH,OXTEUR,OXTUSD,OXTXBT,PAXGETH,PAXGEUR," _
    & "PAXGUSD,PAXGXBT,QTUMETH,QTUMEUR,QTUMUSD,QTUMXBT,SCETH,SCEUR,SCUSD," _
    & "SCXBT,TRXETH,TRXEUR,TRXUSD,TRXXBT,USDCEUR,USDCHF,USDCUSD,USDCUSDT," _
    & "USDTCAD,USDTCHF,USDTEUR,USDTGBP,USDTJPY,USDTZUSD,WAVESETH,WAVESEUR," _
    & "WAVESUSD,WAVESXBT,XBTCHF,XBTDAI,XBTUSDC,XBTUSDT,XDGEUR,XDGUSD,XETCXETH," _
    & "XETCXXBT,XETCZEUR,XETCZUSD,XETHXXBT,XETHXXBT.d,XETHZCAD,XETHZCAD.d," _
    & "XETHZEUR,XETHZEUR.d,XETHZGBP,XETHZGBP.d,XETHZJPY,XETHZJPY.d,XETHZUSD," _
    & "XETHZUSD.d,XLTCXXBT,XLTCZEUR,XLTCZUSD,XMLNXETH,XMLNXXBT,XMLNZEUR," _
    & "XMLNZUSD,XREPXETH,XREPXXBT,XREPZEUR,XREPZUSD,XRPETH,XRPGBP,XRPUSDT," _
    & "XTZETH,XTZEUR,XTZUSD,XTZXBT,XXBTZCAD,XXBTZCAD.d,XXBTZEUR,XXBTZEUR.d," _
    & "XXBTZGBP,XXBTZGBP.d,XXBTZJPY,XXBTZJPY.d,XXBTZUSD,XXBTZUSD.d,XXDGXXBT," _
    & "XXLMXXBT,XXLMZEUR,XXLMZUSD,XXMRXXBT,XXMRZEUR,XXMRZUSD,XXRPXXBT,XXRPZCAD," _
    & "XXRPZEUR,XXRPZJPY,XXRPZUSD,XZECXXBT,XZECZEUR,XZECZUSD,ZEURZUSD,ZGBPZUSD," _
    & "ZUSDZCAD,ZUSDZJPY"

' #############################################################################
' Utils

' #############################################################################

Private Function GetAllAssetPairs(Optional ByVal includeDarkpool As Boolean = False) As String
    Dim method As String
    Dim result As Variant, errJson As Variant
    Dim i As Integer, retVal As String, arrKeys() As Variant, key As String, item As Variant

    ' prepare params
    method = "AssetPairs"
    ' query public API, receive JSON structure
    Set result = API.KrakenQueryPublic(method)
    'Debug.Print method, WebUtils.BeautifyJson(result)

    ' check error
    errJson = result("error")
    If Not IsEmpty(errJson) Then
        GetAllAssetPairs = Empty
        Exit Function
    End If

    ' concat asset pairs
    retVal = ""
    Set result = result("result")
    arrKeys = result.Keys
    For i = 0 To UBound(arrKeys)
        key = arrKeys(i)
        Set item = result(key)

        If Not includeDarkpool And ExcelUtils.EndsWith(key, ".d") Then GoTo NextIteration

        If i > 0 Then
            retVal = retVal & ","
        End If

        retVal = retVal & item("altname")
NextIteration:
    Next

    GetAllAssetPairs = retVal
End Function

Private Sub DoGetTickerInfo()
    Dim method As String
    Dim params As Object
    Dim result As Variant, item As Variant, errJson As Variant
    Dim assetpairs As String
    Dim i As Integer, curOffset As Integer, arrKeys() As Variant, key As String

    Set params = CreateObject("Scripting.Dictionary")

    ' prepare params
    method = "Ticker"

    ' Check if assetpair given
    assetpairs = ActiveSheet.Range(cellInputAssetPairs).Value
    If VBA.Len(assetpairs) > 3 Then
        params.Add "pair", assetpairs
    ElseIf VBA.Len(VBA.Trim(assetpairs)) = 0 Then
        Debug.Print "Warning: No asset pairs provided. Query list of all available pairs ..."
        assetpairs = GetAllAssetPairs(False)
        ActiveSheet.Range(cellInputAssetPairs).Value = assetpairs
        params.Add "pair", assetpairs
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
        If i Mod 25 = 0 Then VBA.DoEvents

        key = arrKeys(i)
        Set item = result(key)

        ActiveSheet.Range("B" & curOffset).Value = key
        ActiveSheet.Range("C" & curOffset).Value = item("a")(0)
        ActiveSheet.Range("D" & curOffset).Value = item("a")(1)
        ActiveSheet.Range("E" & curOffset).Value = item("a")(2)
        ActiveSheet.Range("F" & curOffset).Value = item("b")(0)
        ActiveSheet.Range("G" & curOffset).Value = item("b")(1)
        ActiveSheet.Range("H" & curOffset).Value = item("b")(2)
        ActiveSheet.Range("I" & curOffset).Value = item("c")(0)
        ActiveSheet.Range("J" & curOffset).Value = item("c")(1)
        ActiveSheet.Range("K" & curOffset).Value = item("v")(0)
        ActiveSheet.Range("L" & curOffset).Value = item("v")(1)
        ActiveSheet.Range("M" & curOffset).Value = item("p")(0)
        ActiveSheet.Range("N" & curOffset).Value = item("p")(1)
        ActiveSheet.Range("O" & curOffset).Value = item("t")(0)
        ActiveSheet.Range("P" & curOffset).Value = item("t")(1)
        ActiveSheet.Range("Q" & curOffset).Value = item("l")(0)
        ActiveSheet.Range("R" & curOffset).Value = item("l")(1)
        ActiveSheet.Range("S" & curOffset).Value = item("h")(0)
        ActiveSheet.Range("T" & curOffset).Value = item("h")(1)
        ActiveSheet.Range("U" & curOffset).Value = item("o")

        curOffset = curOffset + 1
    Next
End Sub


' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim cellName As String

    cellName = ExcelUtils.ColumnName(Target.Range.Column) & Target.Range.Row

    If cellLinkButton = cellName Then
        Debug.Print "Get Ticker info ..."
        Application.ScreenUpdating = False
        On Error Resume Next
        DoGetTickerInfo
        Application.ScreenUpdating = True
    End If

End Sub

' #############################################################################



