VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Socket Server"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1590
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   1590
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timTimeOut 
      Interval        =   1000
      Left            =   780
      Top             =   135
   End
   Begin MSWinsockLib.Winsock wsSocket 
      Left            =   285
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================================
'Written By Anoop.M
'website: http://www.geocities.com/anoopj12
'mail   : anoopj12@yahoo.com
'
'IMPORTANT: Read the 'Projects' Section in the associated article
'before seeing this
'
'See the Socket.Cls before seeing this.
'
'============================================================
'FRMLOG.FRM
'The form with the WinSock, to manage transactions
'============================================================
'
'Each time when the class object is created, it will
'create an instance of this form.
'============================================================



'Variables for setting

'ifConnected is the flag we may use for waiting (while connection)
'ConResult is checked by the class when the connectin operation completed
Public ifConnected As Boolean, ConResult

'ifReply is the flag we may use for waiting (while sending data)
'DataToGet is checked by the class after the SendData operation
Public ifReply As Boolean, DataToGet

'TimeOut while sending data
Public TimeOut

'TimeCount is used as a counter
Public TimeCount



Private Sub Form_Load()

'Initialize
TimeCount = 0
ifConnected = False
ifReply = False
DataToGet = ""
End Sub

Private Sub timTimeOut_Timer()
'Timer is enabled and disabled from the class
TimeCount = TimeCount + 1

'Checks the timeout
If TimeCount = TimeOut Then
'Return the error to DataToGet variable
DataToGet = "ERROR 1 Time Out"
    ifReply = True
End If
End Sub

Private Sub wsSocket_Close()
'Returns the error when a socket closes
ConResult = "ERROR 3 Closed"
If DataToGet <> "" Then DataToGet = "ERROR 3 Closed"

'Eliminate waiting
ifConnected = True
ifReply = True
End Sub

Private Sub wsSocket_Connect()
'Eliminate connection wait
ifConnected = True
End Sub


Private Sub wsSocket_DataArrival(ByVal bytesTotal As Long)
'Get data when data comes
wsSocket.GetData DataToGet, vbString

'Eliminate data reply wait
ifReply = True
End Sub

Private Sub wsSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Some error occurred.
ConResult = "ERROR 2 " & Description
ifConnected = True
End Sub
