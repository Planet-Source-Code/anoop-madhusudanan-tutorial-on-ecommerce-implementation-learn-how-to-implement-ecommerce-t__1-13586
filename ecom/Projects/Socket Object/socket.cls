VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'============================================================
'Written By Anoop.M
'website: http://www.geocities.com/anoopj12
'mail   : anoopj12@yahoo.com
'
'============================================================
'IMPORTANT: Read the 'Projects' Section in the associated article
'before seeing this
'============================================================
'
'SOCKET CLASS
'This class is the basic object, that wraps sockets
'
'============================================================

'We need to create a new instance of this form, each time
'the component is initialized
Dim frmNewInst As frmLog



'============================================================
Private Sub Class_Initialize()
'============================================================

'The component is initialized, so need to create a new form
Set frmNewInst = New frmLog

'Load it and hide it
Load frmNewInst
frmNewInst.Hide

'timTimeOut is the Timer we used in the form, to calculate
'transaction timeout
frmNewInst.timTimeOut.Enabled = False

'============================================================
End Sub
'============================================================

'============================================================
Public Function AboutComp() As String
'============================================================
'Just a function. :-). Hope you may not change it.
AboutComp = "Created By Anoop. M, Govindanikethan, Nedumkunnam P.O, Kottayam, Kerala, India" + vbCrLf + "anoopj12@yahoo.com" + vbCrLf + vbCrLf + "Do not use this component commercially, with out the written permission from its author"
'============================================================
End Function
'============================================================

'============================================================
Public Function Connect(HostName, PortNumber) As String
'This method is called first, for establishing a connection.
'
'Hostname: IP of host
'PortNumber: Number of port to connect
'
'Returns: A String. Either a string starting with "ERROR",
'                   or a "NOERROR"
'
'============================================================


On Error Resume Next

'Tries to connect to the destination
frmNewInst.wsSocket.Connect HostName, PortNumber

'Loop to allow connection
Do
'ifConnected becomes true, either when the connection is established, or when an error occurrs.
If frmNewInst.ifConnected = True Then GoTo Skiphere
'Just to do other things
dummy = DoEvents()
Loop

Skiphere:

'The result of connection is stored in the ConResult variable
If Err Or Left(frmNewInst.ConResult, 5) = "ERROR" Then
'We need to return the result to the function, to handle it from ASP.
    Connect = "ERROR"
    frmNewInst.ifConnected = False
    frmNewInst.wsSocket.Close
Else
'We need to return the result to the function, to handle it from ASP
    Connect = "NOERROR"
End If

'============================================================
End Function
'============================================================

'============================================================
Public Function SendData(DataToSend, Optional TimeOut = 4)
'Once the above function returns 'NOERROR', this event
'may be called to send and receive data
'
'DataToSend: The Data To Send, once the connection is established
'Timeout: TimeOut in seconds (for implementing waiting)
'
'Returns: The data recevied after the SendData operation, or
'         a string starting with 'ERROR' if an error occurred.
'
'============================================================


'Sets the Timeout
frmNewInst.TimeOut = TimeOut

'Initialize the timeout counter variable
frmNewInst.TimeCount = 0

'Clears the DataToGet variable, which may hold the data
frmNewInst.DataToGet = ""

'Reset the Timer
frmNewInst.timTimeOut.Enabled = True

'Reset the ifReply flag, which is set when data arrives
frmNewInst.ifReply = False

'Sends the data
frmNewInst.wsSocket.SendData DataToSend

'Wait while getting data
Do
'ifReply is set when data arrives
If frmNewInst.ifReply = True Then GoTo GotData
dummy = DoEvents()
Loop

GotData:
'Got the data
frmNewInst.timTimeOut.Enabled = False

'DataToGet variable in frmNewInst holds the data string retreived
'or the error occurred. Return it to the function

SendData = frmNewInst.DataToGet

'============================================================
End Function
'============================================================


'============================================================
Private Sub Class_Terminate()
'============================================================
On Error Resume Next
'Clean up
frmNewInst.wsSocket.Close
Unload frmNewInst
Set frmNewInst = Nothing
'============================================================
End Sub
'============================================================

