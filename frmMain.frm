VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Telnet Client"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13335
   FillColor       =   &H80000012&
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   13335
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin Project1.telnet telnet 
      Left            =   5370
      Top             =   1590
      _extentx        =   2275
      _extenty        =   688
   End
   Begin RichTextLib.RichTextBox txtTerm 
      Height          =   1545
      Left            =   450
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2430
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   2725
      _Version        =   393217
      BackColor       =   16384
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtCmdLine 
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   405
      TabIndex        =   0
      Top             =   4275
      Width           =   3855
   End
   Begin VB.Label lblStdFont 
      BackColor       =   &H00C0C0FF&
      Caption         =   "std font"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   4500
      TabIndex        =   2
      Top             =   2325
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mnuDisconnect 
      Caption         =   "&Disconnect"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_SCROLLCARET = &HB7
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB

Private Const WM_SETREDRAW = &HB
Private Const WM_USER As Integer = &H400
Private Const EM_GETEVENTMASK As Integer = (WM_USER + 59)
Private Const EM_SETEVENTMASK As Integer = (WM_USER + 69)
Private Const MaxTermLines As Long = 5000

Private Sub Form_Load()
    Me.Caption = DISPLAY_TITLE
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrHandler
    If Me.Height > 0 And Me.Width > 0 Then
        txtTerm.Top = 3 * Screen.TwipsPerPixelY
        txtTerm.Width = Me.ScaleWidth - 8 * Screen.TwipsPerPixelX
        txtTerm.Left = 3 * Screen.TwipsPerPixelX
        txtTerm.Height = Me.ScaleHeight - txtCmdLine.Height - 10 * Screen.TwipsPerPixelY

        txtCmdLine.Left = txtTerm.Left
        txtCmdLine.Width = txtTerm.Width
        txtCmdLine.Top = txtTerm.Top + txtTerm.Height + 3 * Screen.TwipsPerPixelY
    End If
    Exit Sub
ErrHandler:
    'Do nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If telnet.State <> sckClosed Then
        telnet.Disconnect
    End If
End Sub


Private Sub mnuConnect_Click()
    frmConnect.Show vbModal, Me
End Sub


Private Sub mnuDisconnect_Click()
    If telnet.State <> sckClosed Then
        telnet.Disconnect
    End If
    ShowTextInTerm vbCrLf & "Session closed!"
End Sub

Private Sub telnet_Closed()
    ShowTextInTerm vbCrLf & "Remote disconnected!"
End Sub


Private Sub telnet_ConnectTimeout()
    ShowTextInTerm vbCrLf & "Connection time out!"
    If telnet.State <> sckClosed Then
        telnet.Disconnect
    End If
End Sub

Private Sub telnet_Error(ByVal Description As String)
    ShowTextInTerm vbCrLf & Description
End Sub


Public Sub ShowTextInTerm(ByRef ShowText As String, Optional ByVal DoClearFirst As Boolean = False)
    Dim CurLines As Long
    Dim DelOffset As Long
    Dim aryLines() As String
    Dim ShowLines As Long
    Dim OverLines As Long
    Dim CursorRewindLen As Long
    If ShowText = "" Then Exit Sub
    On Error Resume Next
    
    Dim iEventMask As Long
    iEventMask = SendMessage(txtTerm.hwnd, EM_GETEVENTMASK, 0, 0&)
    Call SendMessage(txtTerm.hwnd, WM_SETREDRAW, 0, 0&)

    'cacluate the display rows
    aryLines = Split(ShowText, vbCrLf)

    ShowLines = UBound(aryLines) '+1 -1 (if vbcrlf then has a new line, so minus 1)
    
    'caculate current lines in txtTerm
    CurLines = SendMessage(txtTerm.hwnd, EM_GETLINECOUNT, 0, 0&)
    
    OverLines = CurLines + ShowLines - MaxTermLines
    
    If (txtTerm.Text = "") Or _
        (OverLines >= MaxTermLines) _
        Then DoClearFirst = True
    
    If Not DoClearFirst Then
        'EM_LINEINDEX, first row equal to 1
        DelOffset = SendMessage(txtTerm.hwnd, EM_LINEINDEX, OverLines, 0)
        If DelOffset < 0 Then 'impossible, but...
            DoClearFirst = True
        End If
    End If
    
    
    If DoClearFirst Then
        txtTerm.Text = ""
    Else
        If OverLines > 0 Then
            txtTerm.SelStart = 0
            txtTerm.SelLength = DelOffset
            txtTerm.SelText = ""
        End If
    End If
    
    'append
    txtTerm.SelStart = Len(txtTerm.Text)
    txtTerm.SelLength = 0
    
    txtTerm.SelFontName = lblStdFont.FontName
    txtTerm.SelFontSize = lblStdFont.FontSize
    txtTerm.SelBold = lblStdFont.FontBold
    txtTerm.SelColor = vbWhite
    txtTerm.SelText = ShowText
    
    
    'restore sel properties
    txtTerm.SelLength = 0
    txtTerm.SelStart = Len(txtTerm.Text)
    
   
  
    '###Allow texbox to repaint
    Call SendMessage(txtTerm.hwnd, EM_SETEVENTMASK, 0, iEventMask)
    Call SendMessage(txtTerm.hwnd, WM_SETREDRAW, 1, 0&)
    Call SendMessage(txtTerm.hwnd, EM_SCROLLCARET, 0, ByVal 0)
    txtTerm.Refresh
End Sub


Private Sub telnet_Receive(ByVal DispData As String)
    On Error GoTo ErrHandler

    ShowTextInTerm DispData
      
    Exit Sub
ErrHandler:
    ShowTextInTerm vbCrLf & "Error occurred!" & Err.Description
End Sub

Private Sub txtCmdLine_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cmdline As String
    Dim RetVal As Long
    Select Case KeyCode
    Case 13
        cmdline = Trim(txtCmdLine.Text)
        telnet.SendData txtCmdLine.Text & Chr(13)
        txtCmdLine.Text = ""
    End Select
    
End Sub

Private Sub txtCmdLine_KeyPress(KeyAscii As Integer)
    'disable beep
    Select Case KeyAscii
    Case vbKeyReturn
        'do something
        KeyAscii = 0
    End Select
End Sub


Public Sub ConnectToHost(Host As String, Port As Integer)
    ShowTextInTerm "Connecting to " & Host, True
    telnet.Connect Host, Port
End Sub


