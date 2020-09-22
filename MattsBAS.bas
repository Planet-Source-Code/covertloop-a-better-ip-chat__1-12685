Attribute VB_Name = "MattsBAS"
Public a As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) _
  As Long
Public lShowCursor As Long
Public lRet As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_ALIAS = &H10000
Public Const SND_FILENAME = &H20000
Public Const SND_RESOURCE = &H40004
Public Const SND_ALIAS_ID = &H110000
Public Const SND_ALIAS_START = 0
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_VALID = &H1F
Public Const SND_NOWAIT = &H2000
Public Const SND_VALIDFLAGS = &H17201F

Public Const SND_RESERVED = &HFF000000
Public Const SND_TYPE_MASK = &H170007

Public RAN As Long


Sub CollideDetection(ObjectHitting As Object, ObjectBeingHit As Object)
a = a + 1
If a = 36 Then a = 0
If ObjectHitting.Left + ObjectHitting.Width > ObjectBeingHit(a).Left Then
    If ObjectHitting.Left < ObjectBeingHit(a).Left + ObjectBeingHit(a).Width Then
        If ObjectHitting.Top < ObjectBeingHit(a).Top + ObjectBeingHit(a).Height Then
            If ObjectHitting.Top + ObjectHitting.Height > ObjectBeingHit(a).Top Then
            MsgBox "DOOMED!"
        End If
        End If
        End If
        End If

End Sub


Public Sub HideMouse()
  Do
    lShowCursor = lShowCursor - 1
    lRet = ShowCursor(False)
  Loop Until lRet < 0
End Sub

Sub LoadTextToList(Filename As String, lst As ListBox)
Dim aLine As String
Open Filename For Input As 1
Do Until EOF(1)
    Line Input #1, aLine
    If Trim$(aLine) <> "" Then lst.AddItem aLine
Loop
Close 1
End Sub


Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub


Sub ParseList(StringToRemove As String, lst As ListBox)
Dim X As Integer
Dim Search, Where
lst.ListIndex = -1
On Error GoTo erh
Do
Start:
lst.ListIndex = lst.ListIndex + 1
Search = StringToRemove
Where = InStr(lst.Text, StringToRemove)
If Where Then GoTo Start
lst.RemoveItem lst.ListIndex
Loop Until lst.ListIndex = lst.ListCount - 1
erh:
Exit Sub
End Sub

Sub playwav(File)
SoundName$ = File
SoundFlags& = &H20000 Or &H1
snd& = sndPlaySound(SoundName$, SoundFlags&)
End Sub
Sub AddToList(lst As ListBox, Str As String)
lst.AddItem Str
End Sub


Sub CenterForm(Frm As Form)
Frm.Top = Screen.Height / 2 - Frm.Height / 2
Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub


Sub CMDShowOpen(cmd As Control, DiagTitle As String, Filename As String, Directory As String, Filter As String)
cmd.DialogTitle = DiagTitle
cmd.Filename = Filename
cmd.InitDir = Directory
cmd.Filter = Filter
cmd.ShowOpen
End Sub

Sub CMDShowSave(cmd As Control, DiagTitle As String, Filename As String, Directory As String, Filter As String)
cmd.DialogTitle = DiagTitle
cmd.Filename = Filename
cmd.InitDir = Directory
cmd.Filter = Filter
cmd.ShowSave
End Sub

Sub CopyFile(SrcFile As String, DestFile As String)
Dim sourcefile As String
Dim destinationfile As String
sourcefile = SrcFile
destinationfile = DestFile
FileCopy sourcefile, destinationfile
End Sub

Sub FullScreen(Frm As Form)
Frm.Move 0, 0, Screen.Width, Screen.Height
End Sub


Sub HighlightMe(Txt As TextBox)
Txt.SelStart = 0
Txt.SelLength = Len(Txt.Text)
End Sub

Sub ImageToClip(ImgBox As Image)
Clipboard.Clear
Clipboard.SetData ImgBox.Picture, vbCFBitmap
End Sub

Sub LoadNewControl(Ctrl As Control, Left As Integer, Top As Integer, Indx As Integer, Visibility As Boolean)
Load Ctrl(Indx)
Ctrl(Indx).Left = Left
Ctrl(Indx).Top = Top
Ctrl(Indx).Visible = Visibility
End Sub

Sub LoadTXTtoList(Filename As String, lst As ListBox)
Dim aLine As String
Open Filename For Input As 1
Do Until EOF(1)
    Line Input #1, aLine
    If Trim$(aLine) <> "" Then lst.AddItem aLine
Loop
Close 1
End Sub

Sub MB(message As String)
MsgBox message
End Sub

Sub MenuBarLine(Fram As Frame)
Fram.Width = Screen.Width + 100
Fram.Move -50, 0
End Sub

Sub MPpause(MP As Control)
MP.Pause
End Sub

Sub MPplay(MP As Control)
MP.play

End Sub

Sub MPstop(MP As Control)
MP.Stop
End Sub

Sub MSCOMMdial(MSc As Control, ThePort As Integer, NumberToDial As Integer)
On Error Resume Next
MSc.commport = ThePort
MSc.PortOpen = True
PhoneNumber$ = NumberToDial
MSc.Output = "ATD" + PhoneNumber$ + Chr$(13)
End Sub

Sub MSCOMMhangup(MSc As Control)
MSc.PortOpen = False
End Sub

Sub NextListIndex(lst As ListBox)
On Error Resume Next
lst.ListIndex = lst.ListIndex + 1
End Sub

Sub OpenExe(FileNameAndPath As String)
Result = Shell("start.exe " & FileNameAndPath, vbHide)
End Sub


Sub OpenTXT(Filename As String, Txt As TextBox)
    Dim a As String
    Open Filename For Input As 1
    a = Input(LOF(1), 1)
    Close 1
    Txt = a
End Sub

Sub PicImageToClip(PicBox As PictureBox)
Clipboard.SetData PicBox.Picture, vbCFBitmap
End Sub

Sub PicToClip(PicBox As PictureBox)
Clipboard.Clear
Clipboard.SetData PicBox.Picture, vbCFBitmap
End Sub

Sub PrevListIndex(lst As ListBox)
On Error Resume Next
lst.ListIndex = lst.ListIndex - 1
End Sub

Sub RandomIt(Total As Integer)
RAN = 0
Randomize
RAN = Int(Total * Rnd + 1)
'RandomIt (10)
'MsgBox ran

End Sub

Sub RemoveFromList(lst As ListBox, LstIndex As Integer)
lst.RemoveItem lst.ListIndex
End Sub

Sub RetrieveHTML(Inet As Control, url As String, Filename As String)
Dim X
Dim strsource
strsource = Inet.OpenUrl(url)
X = Filename
Open X For Output As #1
Print #1, strsource
Close #1
End Sub

Sub SaveTXT(Filename As String, Txt As TextBox)
X = Filename
Open X For Output As #1
Print #1, Txt.Text
Close #1

End Sub

Sub Scroll2Bottom(Txt As TextBox)
Txt.SelStart = Len(Txt.Text)
End Sub

Sub SearchAndHiLite(TextToSearchFor As String, WhereToSearch As TextBox)
Dim Search, Where
Search = TextToSearchFor
Where = InStr(WhereToSearch.Text, Search)
If Where Then
WhereToSearch.SelStart = Where - 1
WhereToSearch.SelLength = Len(Search)
End If
End Sub

Public Sub ShowMouse()
  Do
    lShowCursor = lShowCursor - 1
    lRet = ShowCursor(True)
  Loop Until lRet >= 0
End Sub

Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Sub StopAll()
Do
DoEvents
Loop
End Sub


Sub WSclose(WS As Control)
WS.Close
End Sub

Sub WSconnect(WS As Control, IP As String, Port As Integer)
WS.Close
WS.Connect IP, Port
End Sub


Sub WSlisten(WS As Control, Port As Integer)
WS.Close
WS.LocalPort = CLng(Port)
WS.Listen
End Sub


Sub WSsend(WS As Control, TextToSend As String)
WS.SendData TextToSend
End Sub


