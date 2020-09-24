VERSION 5.00
Begin VB.Form SnowHelp 
   Caption         =   "Help"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   Icon            =   "Snow Help.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   2820
      Width           =   5115
   End
   Begin VB.TextBox Text1 
      Height          =   2715
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Snow Help.frx":000C
      Top             =   60
      Width           =   5115
   End
End
Attribute VB_Name = "SnowHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
'on ESC exit
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next

'resize contents
Text1.Height = Height - 960
Text1.Width = Width - 200
BtnOK.Top = Height - 830
BtnOK.Width = Width - 200
End Sub

Private Sub BtnOK_Click()
Hide
End Sub
