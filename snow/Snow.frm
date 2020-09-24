VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Snow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snow"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7830
   Icon            =   "Snow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp"
      Filter          =   "*.bmp"
      Flags           =   4
      MaxFileSize     =   19264
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   900
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton BtnLarger 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton BtnSmaller 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Menu MnuDemo 
      Caption         =   "&Demo"
      Begin VB.Menu MnuReset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu MnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPause 
         Caption         =   "&Freeze Snow Movement"
      End
      Begin VB.Menu MnuAllowMouse 
         Caption         =   "&Allow Mouse Interaction"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuCreateSnow 
         Caption         =   "&Create New Snow"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuHighPriority 
         Caption         =   "&Use High Priority"
      End
      Begin VB.Menu MnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFallsFrom 
         Caption         =   "&Snow Falls From"
         Begin VB.Menu MnuSnowFallTop 
            Caption         =   "&Top"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuSnowFallFrame 
            Caption         =   "&Movable Frame"
         End
      End
      Begin VB.Menu MnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSnowscapeMain 
         Caption         =   "&Snowscape"
         Begin VB.Menu MnuSnowscape1 
            Caption         =   "Snowscape &1 (Default)"
         End
         Begin VB.Menu MnuSnowscape2 
            Caption         =   "Snowscape &2 (Blank)"
         End
         Begin VB.Menu MnuSnowscape3 
            Caption         =   "Snowscape &3 (Wild)"
         End
         Begin VB.Menu MnuSep7 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCustScape 
            Caption         =   "Custom &Snowscape . . ."
         End
      End
      Begin VB.Menu MnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSavePic 
         Caption         =   "S&ave Picture . . ."
      End
      Begin VB.Menu MnuSaveSnowscape 
         Caption         =   "Sa&ve Snowscape . . ."
      End
      Begin VB.Menu MnuOptionsDialog 
         Caption         =   "&Options . . ."
      End
      Begin VB.Menu MnuInfo 
         Caption         =   "&Info . . ."
      End
      Begin VB.Menu MnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu MnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Snow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************
'*  Snow Demo 1.0               *
'*  By Wesley Crossman          *
'*  wesley_crossman@yahoo.com   *
'*  Last Updated: Feb. 2001     *
'*******************************************************
'* If you would like to use this source, feel free to  *
'* do so. All I ask is that you send me a copy of your *
'* program if possible. I would really please me to    *
'* know that my note-in-a-bottle went somewhere! :-)   *
'* Also, if you need help with your project, feel free *
'* to ask me. I have tons of time I don't use and      *
'* would be priviledged to help a fellow programmer.   *
'* Besides, I get bored sometimes!                     *
'*******************************************************

DefInt A-Z 'set default to integers

'used to set/get snow pixels
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long

'all for filling circle
Private Declare Function CreatePen Lib "gdi32" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal nXStart As Long, ByVal nYStart As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'delay for people with slow computers
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Why two arrays? It's faster!
Dim MainX(1 To 500) 'snow location array (Vertical Loc)
Dim MainY(1 To 500) 'snow location array (Horizontal Loc)
'stores the starting position of the line
Dim LineX, LineY
'tells the DoEvented loop to exit
Dim ExitNow As Boolean
'holds the current snowscape number
Dim SnowscapeNum

'sets a flake of snow at X, Y
Sub SnowSet(ByVal X, ByVal Y)
'if there is a snow flake in that location, exit
For a = 1 To 500
 If MainX(a) = X And MainY(a) = Y Then Exit Sub
Next
'loop through possible snow flakes
For a = 1 To 500
 If MainX(a) = 0 Then
  MainX(a) = X
  MainY(a) = Y
  s = 1
  Exit For
 End If
Next
'if one was found to place, draw it
If s Then SetPixelV hdc, MainX(a), MainY(a), RGB(255, 255, 255): Refresh
End Sub

'this is used to fill the circle shape
Sub Paint(ByVal X, ByVal Y, ByVal Colr&, Form4Paint As Form)
hPen& = CreatePen(0, 0, Colr)
hBrush& = CreateSolidBrush(Colr)
SelectObject Form4Paint.hdc, hBrush
ExtFloodFill Form4Paint.hdc, X, Y, Colr, FloodFillBorder
DeleteObject hPen
DeleteObject hBrush
End Sub

Sub MovementMiniEngine()
For a = 1 To 500
 'cache main for better performance
 tx = MainX(a)
 ty = MainY(a)
 'check if snow flake if active
 If tx Then
  'use complex technique if friction is not absolute
  If AbsFric Then
   'simple move ahead or stop algorithm
   If GetPixel(hdc, tx, ty + 1) Then
    'if active flake is not in front, stop it (otherwise just don't move)
    '<FAST>
    d = 1
    For b = 1 To 500
     If ty + 1 = MainY(b) And tx = MainX(b) Then d = 0:  Exit For
    Next
    If d Then tx = 0
    '</FAST>
    'for slow computers, replace all code within <Fast> with "tx=0"
   Else
    SetPixelV hdc, tx, ty, 0: ty = ty + 1
   End If
  Else
   'complex "natural" movement algorithm
   If ty > 479 Then
    tx = 0
   ElseIf GetPixel(hdc, tx, ty + 1) = 0 Then
    SetPixelV hdc, tx, ty, 0
    ty = ty + 1
   ElseIf GetPixel(hdc, tx + 1, ty + 1) = 0 And (CBool(Rnd * FrictionLevel) Or NoFric) Then
    SetPixelV hdc, tx, ty, 0
    ty = ty + 1
    tx = tx + 1
   ElseIf GetPixel(hdc, tx - 1, ty + 1) = 0 And (CBool(Rnd * FrictionLevel) Or NoFric) Then
    SetPixelV hdc, tx, ty, 0
    ty = ty + 1
    tx = tx - 1
   Else
    '<FAST>
    d = 1
    For b = 1 To 500
     If ty + 1 = MainY(b) And tx = MainX(b) Then d = 0:  Exit For
    Next
    'if active flake is not in front, stop it (otherwise just don't move)
    If d Then tx = 0
    '</FAST>
    'for slow computers, replace all code within <Fast> with "tx=0"
   End If
  End If
  'draw new position (with varying colors) if snow flake is active
  If tx Then SetPixelV hdc, tx, ty, SnowPalette(Rnd * 80 + 175)
  'return values to location array
  MainX(a) = tx
  MainY(a) = ty
 End If
Next
End Sub

'snow engine #1 (drop snow from ceiling/sky)
Sub SnowEngine1()
'start demo
For Ml = 1 To NumInCycle
 'create new snow
 If CreateNewSnow Then
  For a = 1 To 500
   If MainX(a) = 0 Then
    MainX(a) = Rnd * 518 + 1
    MainY(a) = 1
    Exit For
   End If
  Next
 End If
 'run the snow movement mini-engine
 MovementMiniEngine
 'refresh if cycle is 1 in RefreshFreq
 If Ml Mod RefreshFreq = 0 Then Refresh
Next
End Sub

'snow engine #2 (drop snow from movable frame)
Sub SnowEngine2()
'speed up frame access
With Frame1
 'cache frame dimensions
 t = .Top + .Height
 f = .Left
 c = (f + .Width - f + 1)
End With
'start demo movement loop
For Ml = 1 To NumInCycle
 'create new snow
 If CreateNewSnow Then
  For a = 1 To 500
   If MainX(a) = 0 Then
    MainX(a) = Rnd * c + f
    MainY(a) = t
    Exit For
   End If
  Next
 End If
 'run the snow movement mini-engine
 MovementMiniEngine
 'refresh if cycle is 1 in RefreshFreq
 If Ml Mod RefreshFreq = 0 Then Refresh
Next
End Sub

Private Sub BtnSmaller_Click()
If Frame1.Width > 70 Then Frame1.Width = Frame1.Width - 20
End Sub

Private Sub BtnLarger_Click()
If Frame1.Width < 660 Then Frame1.Width = Frame1.Width + 20
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
'set frame's new location
Source.Left = X - Source.Width \ 2
Source.Top = Y - Source.Height \ 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
With Frame1
 If .Visible Then
  Select Case KeyCode
  Case vbKeyAdd
   If .Width < 660 Then .Width = .Width + 5
  Case vbKeySubtract
   If .Width > 5 Then .Width = .Width - 5
  End Select
 End If
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
With Frame1
 If .Visible Then
  Select Case KeyAscii
  Case vbKey8
   If .Top > -100 Then .Top = .Top - 5
  Case vbKey2
   If .Top < 450 Then .Top = .Top + 5
  Case vbKey6
   If .Left < 600 Then .Left = .Left + 5
  Case vbKey4
   If .Left > -100 Then .Left = .Left - 5
  End Select
 End If
End With
End Sub

Private Sub Form_Load()
'standard presets
NumInCycle = 3 'allow 3 moves per cycle
SnowDel = 12 'delay snow 12 msecs per cycle
AllowMouse = 1 'allow mouse interaction
CreateNewSnow = 1 'continue to create snow
AllowSnowMove = 1 'allow snow movement
RefreshFreq = 1 'refresh screen every move
FrictionLevel = 55 'the odds are 1 in 55 that the snow will not slide
RActive = 1: GActive = 1: BActive = 1 'set all colors to active
RMult = 1: GMult = 1: BMult = 1 'set color multipliers to one

'select snowscape
MnuSnowscape1_Click
'draw design
MnuReset_Click

'custom snow palette
SetCustPalette

'show form for loop
Show

''Speed Test
't# = Timer
'For a = 1 To 1000
' SnowEngine1
'Next
'Debug.Print Timer - t
'Unload Me: Unload SnowOptions: Unload SnowHelp: Exit Sub

'main program loop
Do
 'if not paused, move snow (actually faster than "If Paused = 0")
 If Paused Then
 Else
  If GeneratorMode Then SnowEngine2 Else SnowEngine1
 End If
 'allow normal vb fuctions
 DoEvents
 'delay to return some power to Windows
 If SnowDel Then Sleep SnowDel
Loop Until ExitNow
'set priority to normal if it's at high
If HighPriority Then SetPriorityClass GetCurrentProcess, &H20

'exit
Unload SnowOptions
Unload SnowHelp
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if right click with no hotkeys depressed
If Button = 2 And Shift = 0 Then PopupMenu MnuDemo, 2, X, Y
'if "mouse control" is unchecked then exit
If AllowMouse = 0 Then Exit Sub

'mouse control
If Button = 1 Then
 SnowSet X, Y
ElseIf Button = 2 Then
 Select Case Shift
 Case 1
  Line (X - 1, Y - 1)-(X, Y), &H2222FF, BF
  Refresh
 Case 2
  LineX = X
  LineY = Y
 Case 4
  Circle (X, Y), 20, &H2222FF
  Refresh
 End Select
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if "mouse control" is unchecked then exit
If AllowMouse = 0 Then Exit Sub

'mouse control
If Button = 1 Then
 SnowSet X, Y
ElseIf Button = 2 And Shift = 4 Then
 Circle (X, Y), 20, &H2222FF
 Refresh
ElseIf Button = 2 And Shift = 1 Then
 Line (X - 1, Y - 1)-(X, Y), &H2222FF, BF
 Refresh
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If AllowMouse Then
 If Button = 2 And Shift = 2 Then Line (LineX, LineY)-(X, Y), &H2222FF: Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
ExitNow = 1
End Sub

Private Sub MnuAbout_Click()
MsgBox "The Incredible Snow Demo" & vbCrLf & "Written By Wesley Crossman" & vbCrLf & "wesley_crossman@yahoo.com", , "Snow"
End Sub

Private Sub MnuAllowMouse_Click()
MnuAllowMouse.Checked = MnuAllowMouse.Checked Xor -1
AllowMouse = MnuAllowMouse.Checked
End Sub

Private Sub MnuCreateSnow_Click()
MnuCreateSnow.Checked = MnuCreateSnow.Checked Xor -1
CreateNewSnow = MnuCreateSnow.Checked
End Sub

Private Sub MnuCustScape_Click()
On Error GoTo errhandle
With CommonDialog1
 .DialogTitle = "Open Snowscape"
 .ShowOpen
 If .filename > "" Then Picture = LoadPicture(.filename) Else Exit Sub
End With
Erase MainX, MainY
MnuSnowscape1.Checked = 0
MnuSnowscape2.Checked = 0
MnuSnowscape3.Checked = 0
MnuCustScape.Checked = 1
Cls
SnowscapeNum = 4
Exit Sub

errhandle:
MsgBox Err.Description
End Sub

Private Sub MnuExit_Click()
Unload SnowHelp
Unload Me
End Sub

Private Sub MnuHighPriority_Click()
If MnuHighPriority.Checked = 0 Then
 'check for slow computer
 If GetSystemMetrics(73) Then
  'ask them if they're sure
  If MsgBox("You have a slow computer." & vbCrLf & "If you use high priority mode, your computer might freeze." & vbCrLf & "Are you sure?", vbYesNo) = vbYes Then
   HighPriority = 1
  Else
   HighPriority = 0
  End If
 Else
  HighPriority = 1
 End If
Else
 HighPriority = 0
End If

'save settings
MnuHighPriority.Checked = HighPriority
SetPriorityClass GetCurrentProcess, IIf(HighPriority, &H80, &H20)
End Sub

Private Sub MnuInfo_Click()
For a = 1 To 500
 If MainX(a) Then nm = nm + 1
Next
MsgBox "Number of Snow Particles Active: " & nm
End Sub

Private Sub MnuInstructions_Click()
SnowHelp.Show
End Sub

Private Sub MnuOptionsDialog_Click()
SnowOptions.Show 1, Me
End Sub

Private Sub MnuPause_Click()
MnuPause.Checked = MnuPause.Checked Xor -1
Paused = MnuPause.Checked
End Sub

Private Sub MnuReset_Click()
Select Case SnowscapeNum
Case 1: MnuSnowscape1_Click
Case 2: MnuSnowscape2_Click
Case 3: MnuSnowscape3_Click
Case 4: Cls: Erase MainX, MainY
End Select
End Sub

Private Sub MnuSavePic_Click()
On Error GoTo errhandle
With CommonDialog1
 .DialogTitle = "Save Image"
 .ShowSave
 f$ = .filename
 If f > vbNullString Then SavePicture Image, f
End With
Exit Sub

errhandle:
MsgBox Err.Description
End Sub

Private Sub MnuSaveSnowscape_Click()
On Error GoTo errhandle
With CommonDialog1
 .DialogTitle = "Save Snowscape"
 .ShowSave
 f$ = .filename
 For a = 1 To 500
  SetPixelV hdc, MainX(a), MainY(a), 0
 Next
 If f > vbNullString Then SavePicture Image, .filename
End With
Exit Sub

errhandle:
MsgBox Err.Description
End Sub

Private Sub MnuSnowFallFrame_Click()
GeneratorMode = 1
Frame1.Visible = 1
MnuSnowFallFrame.Checked = 1
MnuSnowFallTop.Checked = 0
End Sub

Private Sub MnuSnowFallTop_Click()
GeneratorMode = 0
Frame1.Visible = 0
MnuSnowFallFrame.Checked = 0
MnuSnowFallTop.Checked = 1
End Sub

Private Sub MnuSnowscape1_Click()
'erase any traces of last session
Erase MainX, MainY
MnuSnowscape1.Checked = 1
MnuSnowscape2.Checked = 0
MnuSnowscape3.Checked = 0
MnuCustScape.Checked = 0
Cls
'remove any picture
Picture = LoadPicture
SnowscapeNum = 1

'draw & fill circle
Circle (230, 250), 50, RGB(30, 220, 0), , , 0.99
Paint 239, 250, RGB(30, 220, 0), Me

'draw snowscape
Line (228, 190)-(232, 220), 0, BF
Line (212, 220)-(248, 260), 0, BF
Line (300, 420)-(380, 480), vbBlue, BF
Line (410, 390)-(457, 350), vbBlue
Line (410, 391)-(457, 351), vbBlue
Line (463, 350)-(510, 390), vbBlue
Line (463, 351)-(510, 391), vbBlue
For a = 380 To 480 Step 20
 Line (a, 250)-(a + 12, 250), QBColor(11)
Next
Line (100, 350)-(150, 370), vbYellow
Line (100, 351)-(150, 371), vbYellow

'draw snow sponges
For X = 46 To 90
 For Y = 135 To 205
  If CInt(Rnd * 3) = 0 Then PSet (X, Y), vbBlue
 Next
Next
For X = 1 To 44
 For Y = 130 To 215
  If CInt(Rnd * 4) = 0 Then PSet (X, Y), vbRed
 Next
Next
End Sub

Private Sub MnuSnowscape2_Click()
MnuSnowscape1.Checked = 0
MnuSnowscape2.Checked = 1
MnuSnowscape3.Checked = 0
MnuCustScape.Checked = 0
Erase MainX, MainY
Cls
'remove any picture
Picture = LoadPicture
SnowscapeNum = 2
End Sub

Private Sub MnuSnowscape3_Click()
Erase MainX, MainY
MnuSnowscape1.Checked = 0
MnuSnowscape2.Checked = 0
MnuSnowscape3.Checked = 1
MnuCustScape.Checked = 0
Cls
'remove any picture
Picture = LoadPicture
SnowscapeNum = 3

For a = 1 To 120 Step 5
 Circle (320, 400 - a), a, &HAAAF01 - a
Next

For a = -30 To 30 Step 5
 Line (120, 400)-(120 + a * 2, 300), &HDADAFF + a
Next

For a = 1 To 50 Step 3
 Circle (150, 150), 50, &HFF, a / 100, Sin(a), Sin(a)
Next
End Sub
