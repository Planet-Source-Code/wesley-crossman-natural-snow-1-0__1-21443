VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SnowOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Snow Options"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "Snow Options.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Snow Color Formula"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   4200
      TabIndex        =   11
      Top             =   240
      Width           =   2595
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   1620
         ScaleHeight     =   615
         ScaleWidth      =   675
         TabIndex        =   24
         ToolTipText     =   "Sample Color"
         Top             =   2220
         Width           =   735
      End
      Begin VB.CheckBox ChkG 
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2700
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.TextBox TxtGMult 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "1"
         Top             =   2340
         Width           =   735
      End
      Begin VB.CheckBox ChkB 
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1620
         TabIndex        =   18
         Top             =   1320
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.TextBox TxtBMult 
         Height          =   285
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   17
         Text            =   "1"
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox ChkR 
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.TextBox TxtRMult 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   13
         Text            =   "1"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Sample Output"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1620
         TabIndex        =   25
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Multiplier"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2100
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Multiplier"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   19
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1620
         TabIndex        =   16
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Multiplier"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancel"
      Height          =   795
      Left            =   5700
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   795
      Left            =   4200
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   767
      _Version        =   327682
      LargeChange     =   10
      Max             =   500
      TickFrequency   =   10
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   767
      _Version        =   327682
      Max             =   100
      TickFrequency   =   5
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   435
      Left            =   240
      TabIndex        =   8
      Top             =   2580
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   767
      _Version        =   327682
      LargeChange     =   1
      Max             =   255
      TickFrequency   =   15
   End
   Begin ComctlLib.Slider Slider4 
      Height          =   435
      Left            =   240
      TabIndex        =   26
      Top             =   3600
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   767
      _Version        =   327682
      LargeChange     =   1
      Min             =   1
      Max             =   100
      SelStart        =   1
      TickFrequency   =   5
      Value           =   1
   End
   Begin VB.Label RLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   315
      Left            =   480
      TabIndex        =   28
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh Once Every n Times"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   3360
      Width           =   2955
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Odds of Snow Not Sliding (1 out of n):"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2340
      Width           =   3435
   End
   Begin VB.Label FLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   315
      Left            =   480
      TabIndex        =   9
      Top             =   3060
      Width           =   1395
   End
   Begin VB.Label NLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   315
      Left            =   480
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Movements to Make in a Cycle:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   3795
   End
   Begin VB.Label ALabel 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Milliseconds to Relinquish to Windows:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3795
   End
End
Attribute VB_Name = "SnowOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'* This code is not well commented. *
'* It's just an interface into the  *
'* snow's options.                  *
'************************************

DefInt A-Z

Sub PreviewColor()
On Error GoTo errhandle
'if any color multiplier text box is empty, put in a selected 0
If TxtRMult = "" Then TxtRMult = "0": TxtRMult.SelLength = 1
If TxtGMult = "" Then TxtGMult = "0": TxtGMult.SelLength = 1
If TxtBMult = "" Then TxtBMult = "0": TxtBMult.SelLength = 1
r! = IIf(ChkR.Value, 150 * TxtRMult, 0)
r = IIf(r > 255, 255, r)
r = IIf(r < 0, 0, r)
g! = IIf(ChkG.Value, 150 * TxtGMult, 0)
g = IIf(g > 255, 255, g)
g = IIf(g < 0, 0, g)
b! = IIf(ChkB.Value, 150 * TxtBMult, 0)
b = IIf(b > 255, 255, b)
b = IIf(b < 0, 0, b)
Picture1.BackColor = RGB(r, g, b)
Exit Sub

errhandle:
Picture1.BackColor = 0
End Sub

Private Sub BtnCancel_Click()
Unload Me
End Sub

Private Sub BtnOK_Click()
On Error Resume Next
SnowDel = Slider1.Value
NumInCycle = Slider2.Value
FrictionLevel = Slider3.Value
RefreshFreq = Slider4.Value
RActive = ChkR.Value
GActive = ChkG.Value
BActive = ChkB.Value
RMult = TxtRMult
GMult = TxtGMult
BMult = TxtBMult

'save snow palette
SetCustPalette

If Slider3.Value = 0 Then
 AbsFric = 1
 NoFric = 0
 FrictionLevel = 0
ElseIf Slider3.Value = 255 Then
 AbsFric = 0
 NoFric = 1
 FrictionLevel = 1
Else
 AbsFric = 0
 NoFric = 0
 FrictionLevel = Slider3.Value
End If
Unload Me
End Sub

Private Sub ChkB_Click()
PreviewColor
End Sub

Private Sub ChkG_Click()
PreviewColor
End Sub

Private Sub ChkR_Click()
PreviewColor
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BtnOK_Click
If KeyAscii = 27 Then BtnCancel_Click
End Sub

Private Sub Form_Load()
If NoFric Then
 Slider3.Value = 255
ElseIf AbsFric Then
 Slider3.Value = 0
Else
 Slider3.Value = FrictionLevel
End If
Slider3_Scroll
ALabel = SnowDel
NLabel = NumInCycle
If NumInCycle > 1 Then Slider4.Max = NumInCycle
RLabel = RefreshFreq
Slider1.Value = SnowDel
Slider2.Value = NumInCycle
Slider4.Value = RefreshFreq
'Abs's are required because boolean comes out as -1
ChkR.Value = Abs(RActive)
ChkG.Value = Abs(GActive)
ChkB.Value = Abs(BActive)
TxtRMult = RMult
TxtGMult = GMult
TxtBMult = BMult
PreviewColor
End Sub

Private Sub Slider1_Scroll()
ALabel = Slider1.Value
End Sub

Private Sub Slider2_Scroll()
NLabel = Slider2.Value
If NLabel > 1 Then
 Slider4.Max = NLabel
 Slider4.Enabled = 1
Else
 Slider4.Enabled = 0
End If
Slider4_Scroll
End Sub

Private Sub Slider3_Scroll()
Select Case Slider3.Value
Case 255
 FLabel = "Always Slides"
Case 0
 FLabel = "Never Slides"
Case Else
 FLabel = Slider3.Value
End Select
End Sub

Private Sub Slider4_Scroll()
RLabel = Slider4.Value
End Sub

Private Sub TxtBMult_Change()
PreviewColor
End Sub

Private Sub TxtGMult_Change()
PreviewColor
End Sub

Private Sub TxtRMult_Change()
PreviewColor
End Sub
