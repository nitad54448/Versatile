VERSION 4.00
Begin VB.Form about 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   1236
   ClientTop       =   1440
   ClientWidth     =   6396
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   7.8
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Height          =   4224
   HelpContextID   =   6000
   Icon            =   "About.frx":0000
   Left            =   1188
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6396
   ShowInTaskbar   =   0   'False
   Top             =   1164
   Width           =   6492
   Begin VB.CommandButton cmdok 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   7.8
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   7.8
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   2895
   End
   Begin VB.Line line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   10
      Index           =   2
      X1              =   600
      X2              =   480
      Y1              =   240
      Y2              =   600
   End
   Begin VB.Line line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Index           =   1
      X1              =   480
      X2              =   360
      Y1              =   240
      Y2              =   600
   End
   Begin VB.Line line1 
      BorderColor     =   &H00004080&
      BorderWidth     =   10
      Index           =   0
      X1              =   360
      X2              =   240
      Y1              =   240
      Y2              =   600
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   120
      X2              =   6240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   4
      X1              =   120
      X2              =   6240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   120
      X2              =   6240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   7.8
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   6135
   End
   Begin VB.Line line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      X1              =   120
      X2              =   6240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":0442
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   7.8
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      UseMnemonic     =   0   'False
      Width           =   6135
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Versatile 1.00"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   1
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "about"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdok_Click()
If Not (apel) Then main_display.Show
Unload Me
End Sub

Private Sub Form_Activate()
about.MousePointer = 11
DoEvents
Load main_display
about.MousePointer = 0
End Sub

Private Sub Form_Load()
'apel este un boolean global,daca este true atunci provine din help
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
lbl(0).Caption = " The author of this program is: " & vbCrLf & vbCrLf & " Dr. N. D. Dragoe " & vbCrLf & " University of Bucharest " & vbCrLf & " Faculty of Chemistry, Blvd. Elisabeta 4-12," & vbCrLf & "Bucharest, ROMANIA"
Module1.licenta = "Licence: 1012-9715-400057 "
'minor version 2 - cuprinde smooth pe cinci puncte
lbl(3).Caption = licenta
DoEvents
about.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoEvents
End Sub

Private Sub lbl_Click(Index As Integer)

End Sub
