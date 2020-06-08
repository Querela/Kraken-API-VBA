Attribute VB_Name = "FileUtils"
Option Explicit

' Inside the VBE, Go to Tools -> References, then Select "Microsoft Scripting Runtime"
' (NOTE: this is required as here we do early binding, (for IntelliSense), not late/lazy binding)

' #############################################################################

Public Const krakenCredentialsFilename As String = "kraken.key"

' #############################################################################

' #############################################################################

Public Function LoadKrakenCredentials(Optional ByVal filename As String = "") As Dictionary
    Dim fio As New FileSystemObject
    Dim tStream As TextStream
    Dim line As String, parts() As String
    Dim sKey, sValue As String
    Dim creds As New Dictionary

    If ExcelUtils.IsStringEmpty(filename) Then
        Debug.Print "Working dir", ActiveWorkbook.path
        filename = ActiveWorkbook.path & "\" & krakenCredentialsFilename
    End If

    ' Open with default options
    ' https://docs.microsoft.com/de-de/office/vba/language/reference/user-interface-help/opentextfile-method
    Set tStream = fio.OpenTextFile(filename)

    With tStream
        Do While Not .AtEndOfStream
            line = .ReadLine

            parts = Split(line, "=", 3)
            sKey = Trim(parts(0))
            sValue = Trim(parts(1))
            creds.Add sKey, sValue
        Loop
        .Close
    End With

    ' Debug.Print "Creds", WebUtils.BeautifyJson(creds)
    Set LoadKrakenCredentials = creds
End Function

' #############################################################################

' #############################################################################

