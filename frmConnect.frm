VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Connect"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3780
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3780
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   218
      TabIndex        =   5
      Top             =   945
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   2130
      TabIndex        =   4
      Top             =   945
      Width           =   1455
   End
   Begin VB.ComboBox cboPort 
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Text            =   "telnet"
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblPort 
      Caption         =   "PORT"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblHost 
      Caption         =   "HOST"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    Dim Port As Integer
    On Error GoTo ErrHandler
    ' Issue "Connect" command
    Select Case UCase(cboPort.Text)
        Case "TELNET"
            Port = 23
        Case "FTP"
            Port = 21
        Case "SMTP"
            Port = 25
        Case "POP3"
            Port = 110
        Case Else
    End Select
    frmMain.ConnectToHost txtHost.Text, Port
    DoEvents
    Unload Me
    frmMain.Show
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, DISPLAY_TITLE
End Sub

