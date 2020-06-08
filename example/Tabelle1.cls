VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' #############################################################################

' Parameters for crypto name price update eventhandlers
Private Const sheetName As String = "Kraken GetCurrentPrice"
Private Const colnameOfCryptoNames As String = "B"

Private Const numberOfCurrencies As Integer = 9
Private Const offsetCurrencyNameTop As Integer = 4
Private Const offsetCurrencyPriceTop As Integer = 4
Private Const colnameCurrencyPrice As String = "C"
Private Const colnameCurrencyName As String = "B"
Private Const colnameCurrencyChange As String = "D"
Private Const cellnameUpdateDatetime As String = "C3"
Private Const cellnameOldDatetime As String = "D3"

' #############################################################################
' Utils

' #############################################################################
' Worker

Private Sub ProcessCryptoNameLink(ByVal Target As Hyperlink)
    Dim name As String
    Dim price, oldPrice, diffPrice As Double

    name = Target.name
    price = API.RetrieveCurrentPrice(name)
    'price = API.RetrieveCurrentPrice2(name)
    
    oldPrice = Target.Range.offset(0, 1).Value
    Target.Range.offset(0, 1).Value = price
    diffPrice = price - oldPrice
    Target.Range.offset(0, 2).Value = diffPrice
End Sub

' #############################################################################

Public Sub UpdatePrices()
    Debug.Print "Update prices ..."
    
    Dim startTime, endTime
    
    Dim name As String
    Dim price As Double
    Dim i As Integer

    Dim detailsSheet As Worksheet
    Dim currencyNameRange As Range
    Dim currencyPriceRange As Range
    Dim c As Range
    Dim oldColorIndex As Integer
    Dim offsetDiff, offsetNameRangeTop, offsetPriceRangeTop As Integer
    
    'Application.ScreenUpdating = False
    startTime = Time()
    
    With Application.Sheets(sheetName)
        For i = 0 To numberOfCurrencies - 1
            VBA.DoEvents

            ' save old color & highlight as "working"
            With .Cells(i + offsetCurrencyPriceTop, colnameCurrencyPrice).Interior
                oldColorIndex = .ColorIndex
                .ColorIndex = 8
            End With
            ' grab name of currency and retrieve price
            name = .Cells(i + offsetCurrencyNameTop, colnameCurrencyName).Value
            price = API.RetrieveCurrentPrice(name)
            ' set new price and reset color back to previous
            ' TODO: maybe set red if price fell or green if price rose?
            With .Cells(i + offsetCurrencyPriceTop, colnameCurrencyPrice)
                .Value = price
                .Interior.ColorIndex = oldColorIndex
            End With
        Next
    End With
    
    endTime = Time
    Debug.Print "Elapsed time: " & Round(((endTime - startTime) * 24 * 60 * 60), 2) & " sec."
    'Application.ScreenUpdating = True
    
    'Set detailsSheet = Application.Sheets(sheetName)
    'Set currencyNameRange = detailsSheet.Range("B2", "B10")
    'Set currencyPriceRange = detailsSheet.Range("C4", "C12")

    'offsetNameRangeTop = currencyNameRange.Row
    'offsetPriceRangeTop = currencyPriceRange.Row
    'offsetDiff = -offsetNameRangeTop + offsetPriceRangeTop
    'Debug.Print offsetNameRangeTop, offsetPriceRangeTop, offsetDiff
    
    'For Each c In currencyNameRange.Cells
    '    Debug.Print c.Value, "src: " & c.Row, "idx: " & (c.Row - offsetNameRangeTop)
    '    currencyPriceRange.Cells(c.Row - offsetNameRangeTop) = "1"
    'Next

    ' Debug.Print RetrieveCurrentPrice("XXBTZEUR")
    ' Debug.Print RetrieveCurrentPrice("XETHZEUR")
    ' Debug.Print RetrieveCurrentPrice("XXRPZEUR")
    ' Debug.Print RetrieveCurrentPrice("XTZEUR")
    ' Debug.Print RetrieveCurrentPrice("PAXGEUR")
    ' Debug.Print RetrieveCurrentPrice("SCEUR")
    ' Debug.Print RetrieveCurrentPrice("LSKEUR")
    ' Debug.Print RetrieveCurrentPrice("XDGEUR")
    ' Debug.Print RetrieveCurrentPrice("USDTEUR")
    Debug.Print "Done."
End Sub

Public Sub UpdatePrices2()
    Debug.Print "Update prices (v2) ..."
    
    Dim startTime, endTime

    Dim name As String
    Dim price, oldPrice, diff As Double
    Dim oldColorIndex As Integer
    Dim i As Integer
    
    'Application.ScreenUpdating = False
    startTime = Time()
    
    With Application.Sheets(sheetName)
        For i = 0 To numberOfCurrencies - 1
            VBA.DoEvents

            ' save old color & highlight as "working"
            With .Cells(i + offsetCurrencyPriceTop, colnameCurrencyPrice).Interior
                oldColorIndex = .ColorIndex
                .ColorIndex = 8
            End With

            ' grab name of currency and retrieve price
            name = .Cells(i + offsetCurrencyNameTop, colnameCurrencyName).Value
            'price = 0
            'On Error Resume Next
            price = API.RetrieveCurrentPrice(name)

            ' grab old price and compute difference to new price
            oldPrice = .Cells(i + offsetCurrencyPriceTop, colnameCurrencyPrice).Value
            diff = price - oldPrice

            ' set new price and reset color back to previous
            ' TODO: maybe set red if price fell or green if price rose?
            With .Cells(i + offsetCurrencyPriceTop, colnameCurrencyPrice)
                .Value = price
                .Interior.ColorIndex = oldColorIndex
            End With
            ' set diff to old price
            .Cells(i + offsetCurrencyPriceTop, colnameCurrencyChange).Value = diff
        Next

        ' set update datetime
        .Range(cellnameOldDatetime).Value = .Range(cellnameUpdateDatetime).Value
        .Range(cellnameUpdateDatetime).Value = Now()
    End With
    
    endTime = Time
    Debug.Print "Elapsed time: " & Round(((endTime - startTime) * 24 * 60 * 60), 2) & " sec."
    'Application.ScreenUpdating = True

    Debug.Print "Done."
End Sub

' #############################################################################
' 'Official' event handlers

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    ' Check correct sheet?
    ' Target.Application.ActiveSheet
    If ActiveSheet.name = sheetName Then
        ' Check if correct column for crypto names
        If ExcelUtils.ColumnName(Target.Range.Column) = colnameOfCryptoNames Then
            Application.ScreenUpdating = False
            On Error Resume Next
            ProcessCryptoNameLink Target
            Application.ScreenUpdating = True
        End If
    End If
End Sub

' #############################################################################

