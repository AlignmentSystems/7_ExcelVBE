VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "msMessageQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Public Function MessageParser(ErrDescription As String, ErrHelpContext As Long, ErrHelpFile As String, _
ErrLastDllError As Long, ErrNumber As Long, ErrSource As String, ErlLineNumber As String) As Boolean
Dim strBuiltString As String
Const sErrDescription As String = "sDesc"
Const sErrHelpContext As String = "HlpC"
Const sErrHelpFile As String = "HlpF"
Const sErrLastDllError As String = "LastDll"
Const sErrNumber As String = "ErrNumber"
Const sErrSource  As String = "Src"
Const sErlLineNumber As String = "Line"
Const sComputerName As String = "ComputerName"
Const sUserName As String = "UserName"
Dim oWSH As IWshRuntimeLibrary.WshNetwork

Dim strComputerName As String
Dim strUserAndDomainName As String

Set oWSH = New IWshRuntimeLibrary.WshNetwork

strComputerName = oWSH.ComputerName
strUserAndDomainName = oWSH.UserDomain & "\" & oWSH.UserName


strBuiltString = sErrDescription & "[" & ErrDescription & "]" & _
sErrHelpContext & "[" & ErrHelpContext & "]" & _
sErrHelpFile & "[" & ErrHelpFile & "]" & _
sErrLastDllError & "[" & ErrLastDllError & "]" & _
sErrNumber & "[" & ErrNumber & "]" & _
sErrSource & "[" & ErrSource & "]" & _
sErlLineNumber & "[" & ErlLineNumber & "]" & _
sComputerName & "[" & strComputerName & "]" & _
sUserName & "[" & strUserAndDomainName & "]"

MessageParser = SendErrorMessageToMQ(strBuiltString)
strBuiltString = ""

End Function

Private Function SendErrorMessageToMQ(Payload As Variant)
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
'Constants
Const strQname As String = "ExcelIndErrorRep"
Const strQtype As String = "private$"
'Variables
Dim oMSMQQueueInformation As MSMQ.MSMQQueueInfo
Dim oMSMQIndividualQueue As MSMQ.MSMQQueue
Dim oMSMQmsg As MSMQ.MSMQMessage
'
Dim strFormatName As String
Dim strComputerName As String

FL.GetTheWorkstationName strComputerName
strFormatName = "DIRECT=OS:" & strComputerName & "\" & strQtype & "\" & strQname

Set oMSMQQueueInformation = New MSMQ.MSMQQueueInfo
Set oMSMQmsg = New MSMQ.MSMQMessage

oMSMQQueueInformation.FormatName = strFormatName
  
Set oMSMQIndividualQueue = oMSMQQueueInformation.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
  
oMSMQmsg.Label = "Test Message2"
oMSMQmsg.Body = Payload
Debug.Print oMSMQmsg.BodyLength

' Send the message and close the queue.
oMSMQmsg.Send oMSMQIndividualQueue

oMSMQIndividualQueue.Close

Exit Function

ErrorHandler:
MsgBox "Error " + Hex(Err.Number) + " was returned." + Chr(13) + Err.Description

End Function

