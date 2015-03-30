Attribute VB_Name = "numberStationMQ"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       28th March 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Const clngNumerator As Long = 100
Const clngDenominator As Long = 0
Dim MessageQueueImplementation As msMessageQueue

Public Sub ExcelIndustrialisationErrorReportingToMSMQ()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       28th March 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Dim lngBrokenAnswer As Long
Dim ErrDescription As String
Dim ErrHelpContext As Long
Dim ErrHelpFile As String
Dim ErrHelpFileLong As String
Dim ErrLastDllError As Long
Dim ErrNumber As Long
Dim ErrSource As String
Dim ErlLineNumber As String

On Error GoTo ErrHandler

1   Debug.Print "TestNumbers: Testing error handling"

2   If 1 = 2 Then
3       Debug.Print "testing error handling"
4       Debug.Print "testing error handling"
5   End If

6   lngBrokenAnswer = clngNumerator / clngDenominator

7   If 2 = 1 Then
8       Debug.Print "testing error handling"
9       Debug.Print "testing error handling"
10  End If

Exit Sub
ErrHandler:

With Err
    ErrDescription = .Description
    ErrHelpContext = .HelpContext
    ErrHelpFile = .HelpFile
    ErrLastDllError = .LastDllError
    ErrNumber = .Number
    ErrSource = .Source
    ErlLineNumber = Erl
End With
Err.Clear

Set MessageQueueImplementation = New msMessageQueue

If MessageQueueImplementation.MessageParser(ErrDescription, ErrHelpContext, ErrHelpFileLong, ErrLastDllError, _
    ErrNumber, ErrSource, ErlLineNumber) Then
    Debug.Print "Error reported to queue"
End If
'Debug.Print "[TestNumbers2]Number=" & Err.Number & VBA.vbCrLf & "Description=" & Err.Description & VBA.vbCrLf & "LineOfCode=" & Erl
Err.Clear
On Error GoTo 0

End Sub
