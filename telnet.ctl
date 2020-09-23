VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl telnet 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Left            =   720
      Top             =   720
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackColor       =   &H0000C000&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "telnet"
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
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "telnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Closed()
Public Event Receive(ByVal DispData As String)
Public Event Connected()
Public Event Error(ByVal Description As String)
Public Event ConnectTimeout()


'***Telnet Options***
Private Const OPT_BIN = 0 ' Binary Transmission
Private Const OPT_ECHO = 1 ' Echo
Private Const OPT_RECN = 2 ' Reconnection
Private Const OPT_SUPP = 3 ' Suppress Go Ahead
Private Const OPT_APRX = 4 ' Approx Message Size Negotiation
Private Const OPT_STAT = 5 ' Status
Private Const OPT_TIM = 6 ' Timing Mark
Private Const OPT_REM = 7 ' Remote Controlled Trans and Echo
Private Const OPT_OLW = 8 ' Output Line Width
Private Const OPT_OPS = 9 ' Output Page Size
Private Const OPT_OCRD = 10 ' Output Carriage-Return Disposition
Private Const OPT_OHT = 11 ' Output Horizontal Tabstops
Private Const OPT_OHTD = 12 ' Output Horizontal Tab Disposition
Private Const OPT_OFD = 13 ' Output Formfeed Disposition
Private Const OPT_OVT = 14 ' Output Vertical Tabstops
Private Const OPT_OVTD = 15 ' Output Vertical Tab Disposition
Private Const OPT_OLD = 16 ' Output Linefeed Disposition
Private Const OPT_EXT = 17 ' Extended ASCII
Private Const OPT_LOGO = 18 ' Logout
Private Const OPT_BYTE = 19 ' Byte Macro
Private Const OPT_DATA = 20 ' Data Entry Terminal
Private Const OPT_SUP = 21 ' SUPDUP
Private Const OPT_SUPO = 22 ' SUPDUP Output
Private Const OPT_SNDL = 23 ' Send Location
Private Const OPT_TERM = 24 ' Terminal Type
Private Const OPT_EOR = 25 ' End of Record
Private Const OPT_TACACS = 26 ' TACACS User Identification
Private Const OPT_OM = 27 ' Output Marking
Private Const OPT_TLN = 28 ' Terminal Location Number
Private Const OPT_3270 = 29 ' Telnet 3270 Regime
Private Const OPT_X3 = 30 ' X.3 PAD
Private Const OPT_NAWS = 31 ' Negotiate About Window Size
Private Const OPT_TS = 32 ' Terminal Speed
Private Const OPT_RFC = 33 ' Remote Flow Control
Private Const OPT_LINE = 34 ' Linemode
Private Const OPT_XDL = 35 ' X Display Location
Private Const OPT_ENVIR = 36 ' Telnet Environment Option
Private Const OPT_AUTH = 37 ' Telnet Authentication Option
Private Const OPT_NENVIR = 39 ' Telnet Environment Option
'***End of Telnet Options***


'***Telnet Command: 240~255***
Private Const CMD_SE = 240 'end sub negotiation
Private Const CMD_NOP = 241 'nop
Private Const CMD_DM = 242 'data mark--for connect. cleaning
Private Const CMD_BRK = 243 'break
Private Const CMD_IP = 244 'interrupt process--permanently
Private Const CMD_AO = 245 'abort output--but let prog finish
Private Const CMD_AYT = 246 'are you there
Private Const CMD_EC = 247 'erase the current character
Private Const CMD_EL = 248 'erase the current line
Private Const CMD_GA = 249 'you may reverse the line
Private Const CMD_SB = 250 'interpret as subnegotiation
Private Const CMD_WILL = 251 'I will use option
Private Const CMD_WONT = 252 'I won't use option
Private Const CMD_DO = 253 'please, you use option
Private Const CMD_DONT = 254 'you are not to use option
Private Const CMD_IAC = 255 'Interpret Aa Command:
'***End of Telnet Command***

Private Const MOD_INIT = 0
Private Const MOD_IAC_START = 1
Private Const MOD_DONT = 2
Private Const MOD_DO = 3
Private Const MOD_WONT = 4
Private Const MOD_WILL = 5
Private Const MOD_SB = 6
Private Const MOD_WAITFOR_SE = 7

Private DispBuf() As Byte 'Cannot declare Public here
Private DispData As String

Private Sub tmrConnect_Timer()
    RaiseEvent ConnectTimeout
End Sub

Private Sub UserControl_Initialize()
    Width = Label1.Width
    Height = Label1.Height
End Sub
Public Function GetDispData(ByRef RecvData() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim SB_Option As Byte
    Dim CurMod As Integer
    Dim ReplyBuf As String
    Dim strReplyCmd As String
    Dim byteReplyCmd() As Byte
    Dim sep As String
    Dim b As Byte
    Dim u As Long
    
    
    On Error GoTo ErrHandler
    
    'IAC --> WILL/DO/WONT/DONT --> OptionID
    'IAC --> SB --> options data --> SE

    CurMod = MOD_INIT
    u = UBound(RecvData)
    ReDim DispBuf(u)
    j = 0 'Index of DispBuf
    For i = 0 To u 'one byte by byte
        b = RecvData(i)
        If b = 0 Then
            'discard NULL value
            'do nothing
        Else
            Select Case CurMod
            Case MOD_INIT
                If b = CMD_IAC Then '255
                    CurMod = MOD_IAC_START
                Else
                    DispBuf(j) = RecvData(i)
                    j = j + 1
                End If
            Case MOD_IAC_START
                Select Case b
                    Case CMD_DONT
                        CurMod = MOD_DONT
                    Case CMD_DO
                        CurMod = MOD_DO
                    Case CMD_WONT
                        CurMod = MOD_WONT
                    Case CMD_WILL
                        CurMod = MOD_WILL
                    Case CMD_SB 'wait for a SE to end
                        CurMod = MOD_SB
                End Select
            Case MOD_DO
                'Server Request for option
                Select Case RecvData(i)
                    Case OPT_BIN '0 Binary Transmission
                        ReplyBuf = CMD_IAC & "," & CMD_WILL & "," & RecvData(i)
                    Case OPT_SUPP '3 Suppress Go Ahead
                        ReplyBuf = CMD_IAC & "," & CMD_WILL & "," & RecvData(i)
                    Case OPT_TERM '24 Terminal Type
                        ReplyBuf = CMD_IAC & "," & CMD_WILL & "," & RecvData(i)
                    Case OPT_NAWS '31 Negotiate About Window Size (132 columns, 0 rows will receive all data in one time)
                        ReplyBuf = CMD_IAC & "," & CMD_WILL & "," & RecvData(i) & "," & _
                            CMD_IAC & "," & CMD_SB & "," & OPT_NAWS & "," & 0 & "," & 132 & "," & 0 & "," & 0 & "," & CMD_IAC & "," & CMD_SE
                    Case Else
                        ReplyBuf = CMD_IAC & "," & CMD_WONT & "," & RecvData(i)
                End Select
                CurMod = MOD_INIT
            Case MOD_WILL
                Select Case RecvData(i)
                    Case OPT_BIN
                        ReplyBuf = CMD_IAC & "," & CMD_DO & "," & RecvData(i)
                    Case OPT_ECHO
                        'If Not bInEchoRx Then
                        '    bInEchoRx = True
                            ReplyBuf = CMD_IAC & "," & CMD_DO & "," & RecvData(i)
                        'End If
                    Case OPT_SUPP '3 Suppress Go Ahead
                        ReplyBuf = CMD_IAC & "," & CMD_DO & "," & RecvData(i)
                    Case Else
                        ReplyBuf = CMD_IAC & "," & CMD_DONT & "," & RecvData(i)
                End Select
                CurMod = MOD_INIT
            Case MOD_WONT
                Select Case RecvData(i)
                    Case OPT_ECHO '1
                        'If bInEchoRx Then
                        '    bInEchoRx = False
                            ReplyBuf = CMD_IAC & "," & CMD_DONT & "," & RecvData(i)
                        'End If
                    Case OPT_SUPP
                        ReplyBuf = CMD_IAC & "," & CMD_DONT & "," & RecvData(i)
                End Select
                
                CurMod = MOD_INIT
            Case MOD_DONT
                Select Case RecvData(i)
                    Case OPT_ECHO
                        ReplyBuf = CMD_IAC & "," & CMD_WONT & "," & RecvData(i)
                    Case OPT_NAWS
                        ReplyBuf = CMD_IAC & "," & CMD_WONT & "," & RecvData(i)
                End Select
                CurMod = MOD_INIT
            Case MOD_SB 'Indicates that what follows is subnegotiation of the indicated option.
                Select Case RecvData(i)
                    Case OPT_TERM
                        SB_Option = OPT_TERM
                        CurMod = MOD_WAITFOR_SE
                    Case OPT_NENVIR
                        SB_Option = OPT_NENVIR
                        CurMod = MOD_WAITFOR_SE
                    Case Else
                        CurMod = MOD_INIT
                End Select
            Case MOD_WAITFOR_SE
                If RecvData(i) = CMD_SE Then
                    If SB_Option = OPT_TERM Then
                        'IAC SB TERMINAL-TYPE IS ... IAC SE (ex:IAC SB Terminal-Type IS ANSI IAC SE)
                        'The code for IS is 0
                        ReplyBuf = CMD_IAC & "," & CMD_SB & "," & OPT_TERM & "," & 0 & "," & _
                            Asc("V") & "," & Asc("T") & "," & Asc("1") & "," & Asc("0") & "," & Asc("0") & _
                            "," & CMD_IAC & "," & CMD_SE
                    End If
                    CurMod = MOD_INIT
                End If
            End Select
            
            If ReplyBuf <> "" Then
                'Á×§K²Ä¤@­Ó³r¸¹
                If strReplyCmd = "" Then
                    sep = ""
                Else
                    sep = ","
                End If
                strReplyCmd = strReplyCmd & sep & ReplyBuf
                ReplyBuf = ""
            End If
    
        End If
    Next
        
    If j > 0 Then 'there is data
        ReDim Preserve DispBuf(j - 1)
        DispData = StrConv(DispBuf, vbUnicode)
        'Unix to Dos
        Dim dosbuf As String
        dosbuf = Replace(DispData, vbCrLf, Chr(10))
        dosbuf = Replace(dosbuf, Chr(13), "")
        dosbuf = Replace(dosbuf, Chr(10), vbCrLf)
        DispData = dosbuf
    Else
        DispData = ""
    End If
    GetDispData = j 'lenght of data
    'convert it to byte()
    Dim buf
    If strReplyCmd <> "" Then
        buf = Split(strReplyCmd, ",")
        ReDim byteReplyCmd(UBound(buf)) As Byte
        For i = 0 To UBound(buf)
            byteReplyCmd(i) = Val(buf(i))
        Next
        Winsock1.SendData byteReplyCmd 'auto respond TELNET CMD
    End If
    
    
    
    Exit Function
ErrHandler:
    RaiseEvent Error("GetDispData Error!" & vbCrLf & Err.Description)
End Function


Public Sub Connect(Host As String, Port As Integer, Optional Timeout As Long = 30)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    tmrConnect.Interval = CLng(Timeout * 1000)
    tmrConnect.Enabled = True
    Winsock1.Connect Host, Port
End Sub
Public Sub Disconnect()
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
End Sub
Private Sub Winsock1_Close()
'this is remote(telnet server) triggered event
    RaiseEvent Closed
End Sub

Private Sub Winsock1_Connect()
    'disable connection time to count down
    tmrConnect.Enabled = False
    tmrConnect.Interval = 0
    RaiseEvent Connected
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RecvBuf() As Byte
    Dim P As Long
    Dim LastLine As String
    Dim L As Long
    Dim sLine As String
    
    On Error GoTo ErrHandler
    
    Winsock1.GetData RecvBuf, vbByte + vbArray
    
    L = GetDispData(RecvBuf)
    
    '***there is data
    If L > 0 Then
        RaiseEvent Receive(DispData)
    End If
    Exit Sub
ErrHandler:
    RaiseEvent Error("Error occurred!" & Err.Description)
End Sub

Public Function State() As Integer
    State = Winsock1.State
End Function
Public Sub SendData(Data)
    If Winsock1.State = sckConnected Then
        Winsock1.SendData Data
    End If
End Sub

