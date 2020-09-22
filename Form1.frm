VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InterChat"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   2400
      Width           =   4515
      Begin SHDocVwCtl.WebBrowser web 
         Height          =   2695
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4440
         ExtentX         =   7832
         ExtentY         =   4754
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4048
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Connection"
      TabPicture(0)   =   "Form1.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Chat"
      TabPicture(1)   =   "Form1.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Console"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Sender"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SendButton"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Transfer"
      TabPicture(2)   =   "Form1.frx":0D02
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Image1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2895
         Begin VB.OptionButton Option3 
            Caption         =   "I Would Like To Send A File"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton Option4 
            Caption         =   "I Would Like To Send A Picture"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.CommandButton SendButton 
         Caption         =   "Send"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71400
         TabIndex        =   12
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox Sender 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74880
         TabIndex        =   11
         Top             =   1860
         Width           =   3375
      End
      Begin VB.TextBox Console 
         Enabled         =   0   'False
         Height          =   1335
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   400
         Width           =   4215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "I Will Act As The Guest For This Connection Session"
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         ToolTipText     =   " You Will Connect To The Host Of This Session "
         Top             =   900
         Width           =   4095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "I Will Act As The Host For This Connection Session"
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         ToolTipText     =   " You Will Host This Session.  The Guest Will Connect To You "
         Top             =   540
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Host"
         Height          =   735
         Left            =   -74760
         TabIndex        =   5
         Top             =   1260
         Width           =   4095
         Begin VB.CommandButton Command1 
            Caption         =   "Await Connection From Guest"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            ToolTipText     =   " Await Connection From The Guest "
            Top             =   300
            Width           =   3615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Send File"
         Enabled         =   0   'False
         Height          =   975
         Left            =   180
         TabIndex        =   18
         Top             =   1080
         Width           =   4095
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   2535
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2760
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2760
            TabIndex        =   19
            Top             =   600
            Width           =   1215
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   2040
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label2 
            Caption         =   "File Location && Name"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   1755
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Send Picture"
         Enabled         =   0   'False
         Height          =   975
         Left            =   180
         TabIndex        =   13
         Top             =   1080
         Width           =   4095
         Begin MSComDlg.CommonDialog cmd 
            Left            =   2040
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2760
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2760
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Picture Location && Name"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   300
            Width           =   1755
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Guest"
         Height          =   735
         Left            =   -74760
         TabIndex        =   7
         Top             =   1260
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Text            =   "000.000.000.000"
            ToolTipText     =   " The Internet Protocol (IP) Number Of The Host To Connect To "
            Top             =   280
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Connect To Host"
            Height          =   255
            Left            =   2040
            TabIndex        =   8
            ToolTipText     =   " Click Here To Connect To The Host "
            Top             =   300
            Width           =   1815
         End
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   420
         Visible         =   0   'False
         Width           =   1145
      End
   End
   Begin MSWinsockLib.Winsock Chat 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Transfer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chat_Close()
Console.Enabled = False
Sender.Enabled = False
SendButton.Enabled = False
Frame5.Enabled = False
Frame4.Enabled = False
Frame3.Enabled = False
End Sub

Private Sub Chat_Connect()
web.Navigate "http://" & Chat.RemoteHostIP
SSTab1.Tab = 1
Frame5.Enabled = True
Frame4.Enabled = True
Frame3.Enabled = True
Console.Enabled = True
Sender.Enabled = True
SendButton.Enabled = True
Sender.SetFocus
End Sub

Private Sub Chat_ConnectionRequest(ByVal requestID As Long)
Chat.Close
Chat.Accept requestID
web.Navigate "http://" & Chat.RemoteHostIP
SSTab1.Tab = 1
Frame5.Enabled = True
Frame4.Enabled = True
Frame3.Enabled = True
Console.Enabled = True
Sender.Enabled = True
SendButton.Enabled = True
Sender.SetFocus
End Sub


Private Sub Chat_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Chat.GetData strData
If strData = "*PICREQUEST*" Then
web.Refresh
Form1.Height = 5565
Exit Sub
End If
Console.Text = Console.Text & vbCrLf & strData
Console.SetFocus
Console.SelStart = Len(Console.Text)
Sender.SetFocus
End Sub

Private Sub Command1_Click()
Chat.Close
Chat.LocalPort = CLng(187)
Chat.Listen
End Sub

Private Sub Command2_Click()
On Error GoTo erh
Chat.Close
Chat.Connect Text1.Text, 187
Exit Sub
erh:
MsgBox "The server is not accepting connection requests at this time."
Chat.Close
Exit Sub
End Sub


Private Sub Command3_Click()
On Error GoTo erh
cmd.DialogTitle = "Picture To Send..."
cmd.Filename = "*.jpg;*.gif"
cmd.InitDir = CurDir
cmd.Filter = "JPG & GIF Files (*.jpg;*.gif)"
cmd.ShowOpen
Dim sFile As String, sShortFile As String * 67
Dim lRet As Long
sFile = cmd.Filename
lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
sFile = Left(sShortFile, lRet)
Text2.Text = sFile
Image1.Picture = LoadPicture(Text2.Text)
erh:
Exit Sub
End Sub

Private Sub Command4_Click()
If Text2.Text = "" Then
MsgBox "Select A Picture To Send."
Command3_Click
Exit Sub
End If
Transfer.Close
Transfer.LocalPort = CLng(80)
Transfer.Listen
Chat.SendData "*PICREQUEST*"

End Sub


Private Sub Command5_Click()
If Text3.Text = "" Then
MsgBox "Select A File To Send."
Command6_Click
Exit Sub
End If
Transfer.Close
Transfer.LocalPort = CLng(80)
Transfer.Listen
Chat.SendData "*PICREQUEST*"

End Sub

Private Sub Command6_Click()
'On Error GoTo erh
cmd.DialogTitle = "File To Send..."
cmd.Filename = "*.zip;*.rar"
cmd.InitDir = CurDir
cmd.Filter = "WinZIP & WinRAR Files (*.zip;*.rar)"
cmd.ShowOpen
Dim sFile As String, sShortFile As String * 67
Dim lRet As Long
sFile = cmd.Filename
lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
sFile = Left(sShortFile, lRet)
Text3.Text = sFile
erh:
Exit Sub
End Sub

Private Sub Form_Load()
Show
Form1.Caption = "InterChat - ( " & CurrentIP(True) & " )"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Chat.Close
Transfer.Close
Unload Me
End
End Sub


Private Sub Form_Terminate()
On Error Resume Next
Chat.Close
Transfer.Close
Unload Me
End
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Chat.Close
Transfer.Close
Unload Me
End
End Sub


Private Sub Option1_Click()
Frame1.Visible = True
Frame2.Visible = False
Command1.SetFocus
End Sub


Private Sub Option2_Click()
Frame2.Visible = True
Frame1.Visible = False
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub


Private Sub Option3_Click()
Image1.Enabled = False
Frame3.Visible = False
Frame4.Visible = True
Command6_Click
End Sub

Private Sub Option4_Click()
Frame3.Visible = True
Frame4.Visible = False
Image1.Picture = LoadPicture()
Image1.Visible = True
Command3_Click
End Sub

Private Sub SendButton_Click()
Dim strData As String
strData = Chat.LocalHostName & ":  " & Sender.Text
Chat.SendData strData
Console.Text = Console.Text & vbCrLf & strData
Console.SetFocus
Console.SelStart = Len(Console.Text)
Sender.Text = ""
Sender.SetFocus
End Sub


Private Sub Sender_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendButton_Click
KeyAscii = 0
Exit Sub
End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 1 Then Sender.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2_Click
KeyAscii = 0
Exit Sub
End If
End Sub


Private Sub Transfer_ConnectionRequest(ByVal requestID As Long)
Transfer.Close
Transfer.Accept requestID
End Sub

Private Sub Transfer_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
If Option4.Value = True Then
f = FreeFile
temp = ""
Open Text2.Text For Binary As #f
  temp = Input(FileLen(Text2.Text), #f)
Close #f
getimg = temp
Transfer.SendData getimg
End If
If Option3.Value = True Then
f = FreeFile
temp = ""
Open Text3.Text For Binary As #f
  temp = Input(FileLen(Text3.Text), #f)
Close #f
getimg = temp
Transfer.SendData getimg
End If
End Sub


Private Sub Transfer_SendComplete()
Transfer.Close
SSTab1.Tab = 1
Chat.SendData "***Transfer Complete"
End Sub


