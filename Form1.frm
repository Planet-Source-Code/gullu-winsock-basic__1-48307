VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Basic"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Send message:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3615
      Begin VB.CommandButton Command3 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "Listen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "If you are connecting click the 'Connect' button else click on the 'Listen' button."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":5C12
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'code started
'if winsock established a connection then close it first
If Winsock1.State <> sckConnected Then Winsock1.Close
Winsock1.Connect "LocalHost", 1000 'the IP address and the port value
End Sub 'code ends

'Winsock Basic was written down by Gullu
'In this example you should understand the basics of Winsock
'There are mainly two options,they are:
'1. Listen option
'2. Connect option
'If you choose the 'Listen' option then you don't have to put the IP address or
'the computer name. You will just wait for the connection
'But in case you are the listener then you have to set your own port
'as same as the connector's.
'The Connection method:
'The connection method is rather simple but a little bit complicated than the
'Listen option. In the connection option, you have to put the IP (or name)
'address of the computer you are connecting to and the port should be the
'same with the listener. The port can be any number, like 1000 or even 5000!
'Example of IP address: 127.0.0.1 (which is default for all computers)
'If you want to know your IP address, just put this code with a winsock control
'----------------------------------------
'Private Sub Form_Load()
'Form1.Caption = Winsock1.LocalIP
'End Sub
'----------------------------------------
'Ready! Set! Go!!!!
Private Sub Command2_Click() 'code started
'Here we will wait for the connector to connect with us
Winsock1.LocalPort = 1000 'we set our localport
Winsock1.Listen 'we told the winsock to listen or wait for the connection
End Sub 'code end
Private Sub Command3_Click() 'code started
'we will send the data by the 'SendData' method of the winsock
Winsock1.SendData Text1.Text 'winsock will send the data of text1
End Sub 'code ends
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID 'we told winsock to accept the connection
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String 'declare data as string (ex - abc,ABC)
Winsock1.GetData data, vbString 'winsock will get the data that was sent
MsgBox data, vbInformation, "Data arrived" 'collects the data and displays
End Sub
'To test this, compile it and open two instances of it
'Click on the Listen button of the one
'Click on the Connect button of the another
'Type something in the text box and it will appear in
'another's msgbox
