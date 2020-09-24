VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Server"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsAccept 
      Left            =   900
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   315
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Exit"
      Height          =   390
      Left            =   3750
      TabIndex        =   1
      Top             =   2790
      Width           =   1005
   End
   Begin VB.ListBox lstMain 
      Height          =   2595
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   4740
   End
End
Attribute VB_Name = "frmServer"
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
'============================================================
'
'frmServer.frm - The Server Form
'============================================================
'
'
'In fact, this one is a very simple program, for testing
'the socket class we developed. We will echo all the data
'that comes to this server, and simply logs the activites.
'
'============================================================
'
'Anyway, here is the logic to write a server.
'
'In any server, you need a listening socket, and ofcourse,
'few other sockets to accept connection. When a client
'requests a connection, the ConnectionRequest event is
'raised, with the requestID parameter. By using the Accept
'method of sockets, you can accept the requestID. Then, the
'connection is completed.
'
'Scroll down to see the code. I'll describe each part step by step
'
'


'============================================================
Private Sub Form_Load()
'============================================================

'Here we are setting the listening socket

'Need to set the port bofore telling it to listen
wsListen.LocalPort = 4000

'Call the listen method.
wsListen.Listen

'Now the socket is listening the 4000th port.
'
'When you write a web server, you have to listen the
'80th port, for accepting the HTTP requests from browsers
'
'For FTP it is 20 and 21 I think..OOPS, if my memory is
'not bad. 20 for control exchange and 21 for data exchange? :-)
'
'============================================================
End Sub
'============================================================


Private Sub wsAccept_Close()
lstMain.AddItem Time & " - Disconnected"
End Sub

'============================================================
Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
'============================================================

'So, someone (I mean, another socket) requested a connection
'through the 4000th port.

'In this case, it may be our socket component. So just accept
'the connection :-)

'Close the socket if it is already open
On Error Resume Next
wsAccept.Close
wsAccept.Accept requestID

lstMain.AddItem Time & " - Connected"

'IMPORTANT: Here, I am writing this server in such a way
'that it can't accept more than one connection at a time
'If you need to do so, just create an array of
'wsAccept.
'
'When a new request comes, load a new array element of
'wsAccept. Unload that element when it closes.

'============================================================
End Sub
'============================================================

'============================================================
Private Sub wsAccept_DataArrival(ByVal bytesTotal As Long)
'============================================================
'When data comes, just echo it

'First get the data
wsAccept.GetData myDat, vbString

'Just add it to listbox
lstMain.AddItem Time & " - Data: " & myDat
'Then just send it back, by adding some comments
wsAccept.SendData "You send <b>" & myDat & "</b> to server. We received it by " & Time

'============================================================
End Sub
'============================================================
