Attribute VB_Name = "ExcelUtils"
Option Explicit

' Inside the VBE, Go to Tools -> References, then Select Microsoft XML, v6.0
' (or whatever your latest is. This will give you access to the XML Object Library.)
' Also add "Microsoft Scripting Runtime"

' #############################################################################

' #############################################################################

Public Function ColumnName(ByVal iColumn As Integer) As String
    ' From: https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-convert-column-number-to-letter-or-letter-to-number
    ColumnName = Split(Cells(1, iColumn).Address, "$")(1)
End Function

' #############################################################################

Public Function EndsWith(str As String, ending As String) As Boolean
    ' From: http://excelrevisited.blogspot.com/2012/06/endswith.html
    Dim endingLen As Integer
    endingLen = Len(ending)
    EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Public Function StartsWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

Public Function IsStringEmpty(ByVal s As String) As Boolean
    IsStringEmpty = (Len(Trim(CStr(s))) = 0)
End Function

' #############################################################################

Public Function TimestampToExcel(ByVal timestamp As Long) As Double
    TimestampToExcel = VBA.DateSerial(1970, 1, 1) + (((CLng(timestamp) / 60) / 60) / 24)
End Function

Public Function TimestampToExcel_2(ByVal timestamp As Long) As Variant
    If timestamp = 0 Then
        ' If 0 then assume it was empty so we do not want to compute anything
        ' Häcky as heck ...
        TimestampToExcel_2 = ""
        Exit Function
    End If

    TimestampToExcel_2 = VBA.DateSerial(1970, 1, 1) + (((CLng(timestamp) / 60) / 60) / 24)
End Function

' #############################################################################

' #############################################################################
