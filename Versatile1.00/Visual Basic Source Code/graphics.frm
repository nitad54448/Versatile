VERSION 4.00
Begin VB.Form graphics 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5505
   ClientLeft      =   1230
   ClientTop       =   1080
   ClientWidth     =   7770
   ForeColor       =   &H00000000&
   Height          =   6195
   HelpContextID   =   3000
   Icon            =   "graphics.frx":0000
   Left            =   1170
   LinkTopic       =   "Form1"
   ScaleHeight     =   1077.203
   ScaleMode       =   0  'User
   ScaleWidth      =   1126.549
   Top             =   450
   Width           =   7890
   Begin VB.Frame style 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Style"
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   6
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "1"
         Top             =   2520
         Width           =   492
      End
      Begin VB.TextBox txtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   6
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "5"
         Top             =   2520
         Width           =   492
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00C0C0C0&
         Caption         =   "X Grid"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   4680
         Value           =   1  'Checked
         Width           =   1212
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Y Grid"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   1
         Left            =   2880
         TabIndex        =   21
         Top             =   4680
         Value           =   1  'Checked
         Width           =   1092
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&OK"
         Height          =   350
         Left            =   5760
         TabIndex        =   22
         Top             =   4200
         Width           =   1212
      End
      Begin VB.TextBox txtr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   1080
         TabIndex        =   14
         Text            =   "300.0"
         Top             =   3600
         Width           =   732
      End
      Begin VB.TextBox txtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   0
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "5"
         Top             =   840
         Width           =   492
      End
      Begin VB.TextBox txtr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   1
         Left            =   1080
         TabIndex        =   15
         Text            =   "800.0"
         Top             =   3960
         Width           =   732
      End
      Begin VB.TextBox txtr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   2
         Left            =   3600
         TabIndex        =   18
         Text            =   "0.0"
         Top             =   3600
         Width           =   732
      End
      Begin VB.TextBox txtr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   3
         Left            =   3600
         TabIndex        =   19
         Text            =   "1.0"
         Top             =   3960
         Width           =   732
      End
      Begin VB.TextBox txtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   1
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "5"
         Top             =   1080
         Width           =   492
      End
      Begin VB.TextBox txtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   2
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "5"
         Top             =   1440
         Width           =   492
      End
      Begin VB.TextBox txtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   3
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "5"
         Top             =   1680
         Width           =   492
      End
      Begin VB.TextBox txtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   4
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "5"
         Top             =   1920
         Width           =   492
      End
      Begin VB.TextBox txtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   5
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "5"
         Top             =   2160
         Width           =   492
      End
      Begin VB.TextBox txtl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   0
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "0"
         Top             =   840
         Width           =   492
      End
      Begin VB.TextBox txtl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   1
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "1"
         Top             =   1080
         Width           =   492
      End
      Begin VB.TextBox txtl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   2
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "1"
         Top             =   1440
         Width           =   492
      End
      Begin VB.TextBox txtl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   3
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "1"
         Top             =   1680
         Width           =   492
      End
      Begin VB.TextBox txtl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   4
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "1"
         Top             =   1920
         Width           =   492
      End
      Begin VB.TextBox txtl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   204
         Index           =   5
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "1"
         Top             =   2160
         Width           =   492
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   16
         Text            =   "X"
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   1
         Left            =   3600
         MaxLength       =   12
         TabIndex        =   20
         Text            =   "Y"
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lablshp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "diamond"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   6
         Left            =   4080
         TabIndex        =   75
         Top             =   2520
         Width           =   1092
      End
      Begin VB.Label lablshp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "square"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   5
         Left            =   4080
         TabIndex        =   74
         Top             =   2160
         Width           =   852
      End
      Begin VB.Label lablshp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "square"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   4
         Left            =   4080
         TabIndex        =   73
         Top             =   1920
         Width           =   972
      End
      Begin VB.Label lablshp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   3
         Left            =   4080
         TabIndex        =   72
         Top             =   1680
         Width           =   732
      End
      Begin VB.Label lablshp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "cross"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   2
         Left            =   4080
         TabIndex        =   71
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label lablshp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "circle"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   1
         Left            =   4080
         TabIndex        =   70
         Top             =   1080
         Width           =   852
      End
      Begin VB.Label lablshp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "circle"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   0
         Left            =   4080
         TabIndex        =   69
         Top             =   840
         Width           =   852
      End
      Begin VB.Label lblshape 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   3480
         TabIndex        =   68
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblshape 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   3480
         TabIndex        =   67
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblshape 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   3480
         TabIndex        =   66
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblshape 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   3480
         TabIndex        =   65
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblshape 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   3480
         TabIndex        =   64
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblshape 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   3480
         TabIndex        =   63
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblshape 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   3480
         TabIndex        =   62
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Shape"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   19
         Left            =   3480
         TabIndex        =   61
         Top             =   600
         Width           =   852
      End
      Begin VB.Label coll 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   6
         Left            =   6120
         TabIndex        =   60
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label colp 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000080FF&
         Height          =   216
         Index           =   6
         Left            =   2520
         TabIndex        =   59
         Top             =   2520
         Width           =   732
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "crta"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   58
         Top             =   2520
         Width           =   1572
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   7080
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   240
         X2              =   7080
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Points"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   1800
         TabIndex        =   57
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lines"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   5400
         TabIndex        =   56
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1800
         X2              =   4920
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5400
         X2              =   6840
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   1800
         TabIndex        =   55
         Top             =   600
         Width           =   732
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   2520
         TabIndex        =   54
         Top             =   600
         Width           =   852
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Experimental"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Simulated"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   52
         Top             =   1080
         Width           =   1452
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Coats Redfern"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   1572
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Flynn Wall"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   1572
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "van Krevelen"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   49
         Top             =   1920
         Width           =   1572
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Urbanovici Segal"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   48
         Top             =   2160
         Width           =   1572
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5400
         TabIndex        =   47
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   6120
         TabIndex        =   46
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X range"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   45
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Y range"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   238
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   44
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Min."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   43
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Max."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   42
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Min."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   2880
         TabIndex        =   41
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Max."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   2880
         TabIndex        =   40
         Top             =   3960
         Width           =   495
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   1800
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   2880
         X2              =   4320
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label colp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   216
         Index           =   0
         Left            =   2520
         TabIndex        =   39
         Top             =   840
         Width           =   732
      End
      Begin VB.Label colp 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   216
         Index           =   1
         Left            =   2520
         TabIndex        =   38
         Top             =   1080
         Width           =   732
      End
      Begin VB.Label colp 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FF00&
         Height          =   216
         Index           =   2
         Left            =   2520
         TabIndex        =   37
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label colp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H008080FF&
         Height          =   216
         Index           =   3
         Left            =   2520
         TabIndex        =   36
         Top             =   1680
         Width           =   732
      End
      Begin VB.Label colp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H008080FF&
         Height          =   216
         Index           =   4
         Left            =   2520
         TabIndex        =   35
         Top             =   1920
         Width           =   732
      End
      Begin VB.Label colp 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000080FF&
         Height          =   216
         Index           =   5
         Left            =   2520
         TabIndex        =   34
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label coll 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   6120
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin VB.Label coll 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   6120
         TabIndex        =   32
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label coll 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   6120
         TabIndex        =   31
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label coll 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   6120
         TabIndex        =   30
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label coll 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   6120
         TabIndex        =   29
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label coll 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   6120
         TabIndex        =   28
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   26
         Top             =   4320
         Width           =   615
      End
   End
   Begin VB.Label xy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   238
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label xy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   238
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Menu mnu_graph 
      Caption         =   "&Graphic"
      HelpContextID   =   3101
      Begin VB.Menu mnu_draw 
         Caption         =   "&Draw "
      End
      Begin VB.Menu mnu_7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_copy 
         Caption         =   "&Copy To"
         Begin VB.Menu copy_printer 
            Caption         =   "&Printer"
         End
         Begin VB.Menu mnu_picture 
            Caption         =   "&BMP file"
         End
         Begin VB.Menu mnu_graph_copy 
            Caption         =   "&Clipboard"
         End
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_export 
         Caption         =   "&Export ASCII"
      End
      Begin VB.Menu mnu___ 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "&Edit "
      HelpContextID   =   3201
      Begin VB.Menu mnu_follow 
         Caption         =   "&Inspect data"
      End
      Begin VB.Menu m_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_style 
         Caption         =   "&Style"
      End
   End
   Begin VB.Menu mnu_view 
      Caption         =   "&View"
      HelpContextID   =   3301
      Begin VB.Menu mnu_view_comp 
         Caption         =   "&Computed data"
         Begin VB.Menu mnu_comp1 
            Caption         =   "Not available"
         End
         Begin VB.Menu mnu_comp2 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_comp3 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_comp4 
            Caption         =   ""
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu_simulated 
         Caption         =   "&Simulated data"
      End
      Begin VB.Menu mnu_view_exp 
         Caption         =   "&Experimental data"
      End
      Begin VB.Menu mnu_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_icar 
         Caption         =   "&ICAR"
         Enabled         =   0   'False
         Begin VB.Menu icar_exp 
            Caption         =   "&Experimental"
            Begin VB.Menu mnu_curve1 
               Caption         =   "&1st (T, alpha) curve"
            End
            Begin VB.Menu mnu_curve2 
               Caption         =   "&2nd (T,alpha) curve"
            End
         End
         Begin VB.Menu icar_calc 
            Caption         =   "&Calculated"
            Begin VB.Menu icarcalc1 
               Caption         =   "&1st (T, alpha) curve"
            End
            Begin VB.Menu icarcalc2 
               Caption         =   "&2nd (T, alpha) curve"
            End
         End
         Begin VB.Menu m_3 
            Caption         =   "-"
         End
         Begin VB.Menu icar_dif 
            Caption         =   "&Difference"
         End
      End
      Begin VB.Menu m_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sim 
         Caption         =   "Simulate &new"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "graphics"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim stitle As String, sxlabel As String, sylabel As String
Dim slabels As String, slegends As String
Dim pointsize(7) As Integer, linesize(7) As Integer
Dim pointcolor(7) As Long, linecolor(7) As Long
Dim unitx As Single, unity As Single

Sub shapepset(dest As Object, x As Single, y As Single, z As Integer, culoare As Long, stil As Integer)
'z este dimensiunea
On Error GoTo handleit
Dim actualdraw As Integer, actualcolor As Long, i As Integer
actualdraw = dest.DrawWidth
actualcolor = dest.ForeColor
dest.ForeColor = culoare
dest.DrawWidth = 1
Select Case stil
Case 1
'cerc
dest.Circle (x, y), z
Case 2
'cerc umplut
For i = z To 1 Step -1
dest.Circle (x, y), i
Next i
Case 3
'cruce
dest.Line (x - z, y)-(x + z, y)
dest.Line (x, y - z)-(x, y + z)
Case 4
'x
dest.Line (x + z, y - z)-(x - z, y + z)
dest.Line (x - z, y - z)-(x + z, y + z)
Case 5
'patrat
dest.Line (x - z, y - z)-(x + z, y - z)
dest.Line (x + z, y - z)-(x + z, y + z)
dest.Line (x + z, y + z)-(x - z, y + z)
dest.Line (x - z, y + z)-(x - z, y - z)
Case 6
'patrat plin
For i = -z To z
dest.Line (x - z, y - i)-(x + z, y - i)
Next i
Case 7
'romb
dest.Line (x, y - z)-(x + z, y)
dest.Line (x + z, y)-(x, y + z)
dest.Line (x, y + z)-(x - z, y)
dest.Line (x - z, y)-(x, y - z)
End Select
dest.DrawWidth = actualdraw
dest.ForeColor = actualcolor
Exit Sub
handleit:
MsgBox ("Unexpected error")
Exit Sub
End Sub

Sub redeseneaza()
'graficul il trasez doar in coordonate de temp, alpha
On Error GoTo handle:
Screen.MousePointer = 11
Dim xg(200) As Double, yg(200) As Double, eroare As Boolean
graphics.Cls
If graphics.chk(0).Value Then graphics.punegridx
If graphics.chk(1).Value Then graphics.punegridy
xstart = CDbl(graphics.txtr(0).Text)
ystart = CDbl(graphics.txtr(2).Text)
xend = CDbl(graphics.txtr(1).Text)
yend = CDbl(graphics.txtr(3).Text)
Call graphics.scrie_labelx(xstart, xend, CStr(graphics.txt(0).Text))
Call graphics.scrie_labely(ystart, yend, CStr(graphics.txt(1).Text))

If data_editor.tabdata.Caption = "CRTA" Then
If mnu_curve1.Checked Then
For i = 1 To ipoints1
xg(i) = dalpha(i): yg(i) = dtempk(i)
Next i
Call desen(ipoints1, xg(), yg(), xstart, xend, ystart, yend, pointsize(1), pointcolor(1), linesize(1), linecolor(1), CInt(lblshape(0).Caption), eroare)
End If
If mnu_curve2.Checked Then
For i = ipoints1 + 2 To ipoints1 + ipoints2 + 1
xg(i - ipoints1 - 1) = dalpha(i): yg(i - ipoints1 - 1) = dtempk(i)
Next i
Call desen(ipoints2, xg(), yg(), xstart, xend, ystart, yend, pointsize(2), pointcolor(2), linesize(2), linecolor(2), CInt(lblshape(1).Caption), eroare)
End If
If icarcalc1.Checked Then
j% = 1
For i = 1 To nrint
xg(i) = dalpha(1) + (i - 1) * (dalpha(ipoints1) - dalpha(1)) / nrint
yg(i) = igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i))
Next i
Call desen(nrint, xg(), yg(), xstart, xend, ystart, yend, pointsize(3), pointcolor(3), linesize(3), linecolor(3), CInt(lblshape(2).Caption), eroare)
End If
If icarcalc2.Checked Then
j% = 2
For i = 1 To nrint
xg(i) = dalpha(ipoints1 + 2) + (i - 1) * (dalpha(ipoints1 + ipoints2 + 1) - dalpha(ipoints1 + 2)) / nrint
yg(i) = igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i))
Next i
Call desen(nrint, xg(), yg(), xstart, xend, ystart, yend, pointsize(4), pointcolor(4), linesize(4), linecolor(4), CInt(lblshape(3).Caption), eroare)
End If

   If mnu_simulated.Checked Then
    For i = 1 To 200
    xg(i) = xgraf(i, 2): yg(i) = ygraf(i, 2)
    Next i
    Call desen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(7), pointcolor(7), linesize(7), linecolor(7), CInt(lblshape(6).Caption), eroare)
    End If

If icar_dif.Checked Then
j% = 1
For i = 1 To ipoints1
xg(i) = dalpha(i)
yg(i) = dtempk(i) - igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i))
Next i
Call desen(ipoints1, xg(), yg(), xstart, xend, ystart, yend, pointsize(5), pointcolor(5), linesize(5), linecolor(5), CInt(lblshape(4).Caption), eroare)
j% = 2
For i = ipoints1 + 2 To ipoints1 + ipoints2 + 1
xg(i - ipoints1 - 1) = dalpha(i)
yg(i - ipoints1 - 1) = dtempk(i) - igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), dalpha(i))
Next i
Call desen(ipoints2, xg(), yg(), xstart, xend, ystart, yend, pointsize(6), pointcolor(6), linesize(6), linecolor(6), CInt(lblshape(5).Caption), eroare)
End If

Else
    If mnu_view_comp.Enabled Then
    
        If mnu_comp1.Checked Then
        For i = 1 To 200
        xg(i) = xgraf(i, 3): yg(i) = ygraf(i, 3)
        Next i
        Call desen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(3), pointcolor(3), linesize(3), linecolor(3), CInt(lblshape(2).Caption), eroare)
        End If
    
        If mnu_comp2.Checked Then
        For i = 1 To 200
        xg(i) = xgraf(i, 4): yg(i) = ygraf(i, 4)
        Next i
        Call desen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(4), pointcolor(4), linesize(4), linecolor(4), CInt(lblshape(3).Caption), eroare)
        End If
        
        If mnu_comp3.Checked Then
        For i = 1 To 200
        xg(i) = xgraf(i, 5): yg(i) = ygraf(i, 5)
        Next i
        Call desen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(5), pointcolor(5), linesize(5), linecolor(5), CInt(lblshape(4).Caption), eroare)
        End If
        
        If mnu_comp4.Checked Then
        For i = 1 To 200
        xg(i) = xgraf(i, 6): yg(i) = ygraf(i, 6)
        Next i
        Call desen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(6), pointcolor(6), linesize(6), linecolor(6), CInt(lblshape(5).Caption), eroare)
        End If
    
  
    
    
    
    End If

    If mnu_view_exp.Checked Then
    Call desen(ipoints, dtempk(), dalpha(), xstart, xend, ystart, yend, pointsize(1), pointcolor(1), linesize(1), linecolor(1), CInt(lblshape(0).Caption), eroare)
    End If

    If mnu_simulated.Checked Then
    For i = 1 To 200
    xg(i) = xgraf(i, 2): yg(i) = ygraf(i, 2)
    Next i
    Call desen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(2), pointcolor(2), linesize(2), linecolor(2), CInt(lblshape(1).Caption), eroare)
    End If

End If
'desenez un frame
graphics.Line (200 - 10, 0 + 10)-(graphics.ScaleWidth - 2 - 10, graphics.ScaleHeight - 100 + 10), 0, B
graphics.Refresh
Screen.MousePointer = 0
Exit Sub
handle:
Screen.MousePointer = 0
Exit Sub
End Sub

Sub scrie_labelx(xstart As Double, xend As Double, unit As String)
pas# = (xend - xstart) / 5#
For i% = 1 To 5
graphics.CurrentY = graphics.ScaleHeight - 80 + 10
graphics.CurrentX = 175 + (i% - 1) * (graphics.ScaleWidth - 200) / 5 - 10
graphics.Print Format$((xstart) + (i% - 1) * pas#, "##0.00")
Next i%
graphics.CurrentY = graphics.ScaleHeight - 80
graphics.CurrentX = graphics.ScaleWidth - 100
graphics.Print unit
End Sub

Sub scrie_labely(ystart As Double, yend As Double, unit As String)
pas# = (yend - ystart) / 4#
For i% = 1 To 4
graphics.CurrentX = 130 - 20
graphics.CurrentY = 10 + graphics.ScaleHeight - 125 - (i% - 1) * (graphics.ScaleHeight - 100) / 4
graphics.Print Format$((ystart) + (i% - 1) * pas#, "##0.00")
Next i%
graphics.CurrentX = 75
graphics.CurrentY = 25
graphics.Print unit
End Sub
Sub punegridx()
On Error GoTo handle
For i% = 1 To 3
graphics.Line (200 - 10, 10 + (graphics.ScaleHeight - 100 - i% * (graphics.ScaleHeight - 100) / 4))-(graphics.ScaleWidth - 10, 10 + (graphics.ScaleHeight - 100 - i% * (graphics.ScaleHeight - 100) / 4)), &HC0C0C0
Next i%
Exit Sub
handle:
Exit Sub
End Sub

Sub punegridy()
On Error GoTo handle
For i% = 1 To 4
graphics.Line (200 + i% * (graphics.ScaleWidth - 200) / 5 - 10, 10 + 0)-(-10 + 200 + i% * (graphics.ScaleWidth - 200) / 5, 10 + graphics.ScaleHeight - 100), &HC0C0C0
Next i%
Exit Sub
handle:
Exit Sub

End Sub

Sub setupgraph()
'aici setez toate enabled sau nu in functie de datele pe care le am ?
End Sub

Sub printdesen(gpuncte As Integer, xg() As Double, yg() As Double, xstart As Double, xend As Double, ystart As Double, yend As Double, point As Integer, colpoint As Long, lsize As Integer, colline As Long, stil As Integer, eroare As Boolean)

On Error GoTo handleit
If lsize > 0 Then
Printer.DrawWidth = lsize + 1
Printer.ForeColor = colline
For i = 1 To gpuncte - 1
If ((xg(i) <= xend) And (xg(i) >= xstart) And (yg(i) <= yend) And (yg(i) >= ystart)) Then
Printer.Line (unitx / 8 + ((xg(i) - xstart) / (xend - xstart)) * (unitx * 3 / 4), unity * 7 / 8 - ((yg(i) - ystart) / (yend - ystart)) * (unity * 3 / 4))-(unitx / 8 + ((xg(i + 1) - xstart) / (xend - xstart)) * (unitx * 3 / 4), unity * 7 / 8 - ((yg(i + 1) - ystart) / (yend - ystart)) * (unity * 3 / 4))
End If
Next i
End If
If point > 0 Then
'Printer.DrawWidth = point * 2 + 1
'Printer.ForeColor = colpoint
For i = 1 To gpuncte
If ((xg(i) <= xend) And (xg(i) >= xstart) And (yg(i) <= yend) And (yg(i) >= ystart)) Then
'Printer.PSet (unity * 7 / 8 - ((yg(i) - ystart) / (yend - ystart)) * (unity * 3 / 4), unitx / 8 + ((xg(i) - xstart) / (xend - xstart)) * (unitx * 3 / 4))
'Printer.PSet (Printer.Width / 8 + (xg(i) - xstart) / (xend - xstart) * Printer.Width * 3 / 4, 7 / 8 * Printer.Height - (yg(i) - ystart) / (yend - ystart) * Printer.Height * 3 / 4), colpoint
Call shapepset(Printer, unitx / 8 + ((xg(i) - xstart) / (xend - xstart)) * (unitx * 3 / 4), unity * 7 / 8 - ((yg(i) - ystart) / (yend - ystart)) * (unity * 3 / 4), 7 * point, colpoint, stil)
End If
Next i
End If
eroare = False
Printer.DrawWidth = 1
Exit Sub
handleit:
eroare = True
Printer.DrawWidth = 1
Exit Sub
End Sub



Sub desen(gpuncte As Integer, xg() As Double, yg() As Double, xstart As Double, xend As Double, ystart As Double, yend As Double, point As Integer, colpoint As Long, lsize As Integer, colline As Long, stil As Integer, eroare As Boolean)
'gpuncte este numarul maxim de puncte
'daca dest este true atunci e la imprimanta
On Error GoTo handleit
If lsize > 0 Then
graphics.DrawWidth = lsize
For i = 1 To gpuncte - 1
If ((xg(i) <= xend) And (xg(i) >= xstart) And (yg(i) <= yend) And (yg(i) >= ystart)) Then
graphics.Line (200 + (graphics.ScaleWidth - 200) * (xg(i) - xstart) / (xend - xstart) - 10, 10 + (graphics.ScaleWidth - 200) - (graphics.ScaleWidth - 200) * (yg(i) - ystart) / (yend - ystart))-(200 + (graphics.ScaleWidth - 200) * (xg(i + 1) - xstart) / (xend - xstart) - 10, 10 + (graphics.ScaleWidth - 200) - (graphics.ScaleWidth - 200) * (yg(i + 1) - ystart) / (yend - ystart)), colline
End If
Next i
End If
If point > 0 Then
'graphics.DrawWidth = point
'graphics.ForeColor = colpoint
For i = 1 To gpuncte
If ((xg(i) <= xend) And (xg(i) >= xstart) And (yg(i) <= yend) And (yg(i) >= ystart)) Then
'graphics.PSet (200 + (graphics.ScaleWidth - 200) * (xg(i) - xstart) / (xend - xstart) - 10, 10 + (graphics.ScaleWidth - 200) - (graphics.ScaleWidth - 200) * (yg(i) - ystart) / (yend - ystart)), colpoint
Call shapepset(graphics, 200 + (graphics.ScaleWidth - 200) * (xg(i) - xstart) / (xend - xstart) - 10, 10 + (graphics.ScaleWidth - 200) - (graphics.ScaleWidth - 200) * (yg(i) - ystart) / (yend - ystart), point, colpoint, stil)
End If
Next i
End If

eroare = False
graphics.DrawWidth = 1
Exit Sub
handleit:
eroare = True
graphics.DrawWidth = 1
Exit Sub
End Sub

Sub cmdok_Click()
On Error GoTo handle
For i% = 0 To 6
If Val(txtp(i%).Text) < 0 Then txtp(i%).Text = "0"
If Val(txtl(i%).Text) < 0 Then txtl(i%).Text = "0"
txtp(i%).Text = CStr(Val(txtp(i%).Text))
txtl(i%).Text = CStr(Val(txtl(i%).Text))
Next i%
If (CDbl(txtr(1).Text) <= CDbl(txtr(0).Text)) Then
MsgBox "Incorect values for the X range. Try again."
Exit Sub
End If
If (CDbl(txtr(3).Text) <= CDbl(txtr(2).Text)) Then
t = MsgBox("Atention :incorect values for the Y range.", vbOKOnly, nume_prog)
Exit Sub
End If

graphics.BackColor = &HFFFFFF
graphics.BackColor = &HFFFFFF
'graphics.Visible = True
For i% = 0 To 6
pointcolor(i% + 1) = colp(i%).BackColor
linecolor(i% + 1) = coll(i%).BackColor
pointsize(i% + 1) = Val(txtp(i%))
linesize(i% + 1) = Val(txtl(i%))
Next i%
xstart = CDbl(txtr(0).Text)
xend = CDbl(txtr(1).Text)
ystart = CDbl(txtr(2).Text)
yend = CDbl(txtr(3).Text)
mnu_graph.Enabled = True
mnu_edit.Enabled = True
mnu_view.Enabled = True
Style.Visible = False
mnu_draw_Click
Exit Sub
handle:
Exit Sub

End Sub

Private Sub coll_Click(Index As Integer)
On Error GoTo handleit
main_display.comdialog1.Flags = &H1&
main_display.comdialog1.CancelError = True
main_display.comdialog1.ShowColor
coll(Index).BackColor = main_display.comdialog1.Color
pointcolor(Index + 1) = coll(Index).BackColor
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub colp_Click(Index As Integer)
On Error GoTo handleit
main_display.comdialog1.Flags = &H1&
main_display.comdialog1.CancelError = True
main_display.comdialog1.ShowColor
colp(Index).BackColor = main_display.comdialog1.Color
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub copy_printer_Click()
On Error GoTo handleit
If mnu_follow.Checked Then mnu_follow_Click
mnu_draw_Click
graphics.Refresh
main_display.comdialog1.Flags = cdlPDReturnDC Or cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoWarning Or cdlPDUseDevModeCopies Or cdlPDAllPages Or cdlPDNoSelection
main_display.comdialog1.ShowPrinter
For i% = 1 To main_display.comdialog1.Copies
Printer.PaintPicture graphics.Image, Printer.Width / 30, Printer.Height / 25
Printer.EndDoc
Next i%
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub Form_Activate()
'If Not (Style.Visible) Then Call mnu_draw_Click
End Sub

Private Sub Form_Load()
Dim i As Integer, eroare As Boolean
On Error GoTo handle
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Me.Caption = nume_prog & " - graphics"
For i = 0 To 6: pointcolor(i + 1) = graphics.colp(i).BackColor: linecolor(i + 1) = graphics.coll(i).BackColor: pointsize(i + 1) = CInt(graphics.txtp(i).Text): linesize(i + 1) = CInt(graphics.txtl(i).Text): Next i

'If gindicator(2) Then
'End If
'If gindicator(3) Then
'End If
'If gindicator(4) Then
'End If
'If gindicator(5) Then
'End If
'If gindicator(6) Then
'End If

mnu_view_exp.Enabled = gindicator(1)
mnu_view_exp.Checked = mnu_view_exp.Enabled
mnu_simulated.Enabled = gindicator(2)

Select Case data_editor.tabdata.Caption

Case "Integral"
If gindicator(1) Then
    txtr(0).Text = CStr(CInt(dtempk(1) - 1))
    txtr(1).Text = CStr(CInt(dtempk(ipoints) + 1))
End If
   
    mnu_comp1.Caption = "Coats-Redfern": mnu_comp1.Visible = True: mnu_comp1.Enabled = gindicator(3)
    mnu_comp2.Caption = "Flynn-Wall": mnu_comp2.Visible = True: mnu_comp2.Enabled = gindicator(4)
    mnu_comp3.Caption = "van Krevelen": mnu_comp3.Visible = True: mnu_comp3.Enabled = gindicator(5)
    mnu_comp4.Caption = "Urbanovici-Segal": mnu_comp4.Visible = True: mnu_comp4.Enabled = gindicator(6)
    lbl(5).Caption = "Coats Redfern": lbl(6).Caption = "Flynn Wall": lbl(7).Caption = "van Krevelen": lbl(8).Caption = "Urbanovici Segal"
For i = 0 To 6
lbl(3 + i).Visible = gindicator(i + 1): txtp(i).Visible = gindicator(i + 1): txtl(i).Visible = gindicator(i + 1): coll(i).Visible = gindicator(i + 1): colp(i).Visible = gindicator(i + 1)
lblshape(i).Visible = gindicator(i + 1)
Next i

Case "Differential"
If gindicator(1) Then
    txtr(0).Text = CStr(CInt(dtempk(1) - 1))
    txtr(1).Text = CStr(CInt(dtempk(ipoints) + 1))
End If
    
    mnu_comp1.Caption = "Achar et al.": mnu_comp1.Visible = True: mnu_comp1.Enabled = gindicator(3)
    mnu_comp2.Caption = "Freeman-Carroll": mnu_comp2.Visible = True: mnu_comp2.Enabled = gindicator(4)
    mnu_comp3.Visible = False
    mnu_comp4.Caption = "Fatu (DTA)": mnu_comp4.Visible = True: mnu_comp4.Enabled = gindicator(6)
lbl(5).Caption = "Achar": lbl(6).Caption = "Freeman Carroll": lbl(7).Caption = "": lbl(8).Caption = "Fatu (DTA)"
For i = 0 To 6
lbl(3 + i).Visible = gindicator(i + 1): txtp(i).Visible = gindicator(i + 1): txtl(i).Visible = gindicator(i + 1): coll(i).Visible = gindicator(i + 1): colp(i).Visible = gindicator(i + 1)
lblshape(i).Visible = gindicator(i + 1)
Next i


Case "Regression"
If gindicator(1) Then
    txtr(0).Text = CStr(CInt(dtempk(1) - 1))
    txtr(1).Text = CStr(CInt(dtempk(ipoints) + 1))
End If
    
    mnu_comp1.Caption = "Not available": mnu_comp1.Visible = True
    mnu_comp2.Visible = False
    mnu_comp3.Visible = False
    mnu_comp4.Visible = False
    lbl(5).Caption = "pseudo-Inverse": lbl(6).Caption = "": lbl(7).Caption = "": lbl(8).Caption = ""

For i = 0 To 6
lbl(3 + i).Visible = gindicator(i + 1): txtp(i).Visible = gindicator(i + 1): txtl(i).Visible = gindicator(i + 1): coll(i).Visible = gindicator(i + 1): colp(i).Visible = gindicator(i + 1)
lblshape(i).Visible = gindicator(i + 1)
Next i


Case "CRTA"
If gindicator(1) Then
    txtr(0).Text = "0.0"
    txtr(1).Text = "1.0"
End If

mnu_view_exp.Enabled = False
mnu_view_comp.Enabled = False
mnu_simulated = False
mnu_sim = True
lbl(3).Caption = "Exp. - 1st curve"
lbl(4).Caption = "Exp. - 2nd curve"
lbl(5).Caption = "Calc. - 1st curve"
lbl(6).Caption = "Calc. - 2nd curve"
lbl(7).Caption = "Diff. - 1st curve"
lbl(8).Caption = "Diff. - 2nd curve"
lbl(9).Caption = "CRTA -Simulated"

For i = 0 To 6
lbl(3 + i).Visible = True: txtp(i).Visible = True: txtl(i).Visible = True: coll(i).Visible = True: colp(i).Visible = True
lblshape(i).Visible = gindicator(i + 1)
Next i
    
    mnu_icar.Enabled = gindicator(3)
    txtr(0).Text = "0.00"
    txtr(1).Text = "1.00"
    txtr(2).Text = CStr(CInt(dtempk(1) - 30))
    txtr(3).Text = CStr(CInt(dtempk(ipoints1) + 10))
End Select

xstart = CDbl(txtr(0).Text): xend = CDbl(txtr(1).Text): ystart = CDbl(txtr(2).Text): yend = CDbl(txtr(3).Text)
Exit Sub
handle:
Exit Sub

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo handle
If mnu_follow.Checked Then
xy(0).Left = x - xy(0).Width
xy(0).Top = Me.ScaleHeight - xy(0).Height
xy(1).Left = 0
xy(1).Top = y


xy(0).Caption = "X=" & Format$(xstart + ((x + 10 - 200) / (graphics.ScaleWidth - 200)) * (xend - xstart), "#####0.00")
xy(1).Caption = "Y=" & Format$(yend - ((y - 10) / (graphics.ScaleHeight - 100)) * (yend - ystart), "####0.00")
DoEvents
End If
Exit Sub
handle:
Exit Sub

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
graphics.PopupMenu mnu_view
End If

End Sub

Private Sub Form_Resize()
On Error GoTo handle
Me.Caption = nume_prog & " - graphics"
If Me.WindowState = 1 Then Me.Caption = " Graphics."
graphics.Cls
'graphics.Height = Me.Height
'graphics.Width = Me.Width
graphics.ScaleMode = 0
graphics.ScaleWidth = 1200
graphics.ScaleHeight = 1100

redeseneaza
Exit Sub
handle:
Exit Sub
End Sub



Private Sub Form_Unload(Cancel As Integer)
gindicator(2) = False
Unload Me
End Sub

Private Sub icar_dif_Click()
icar_dif.Checked = Not (icar_dif.Checked)
End Sub

Private Sub icarcalc1_Click()
icarcalc1.Checked = Not (icarcalc1.Checked)
End Sub

Private Sub icarcalc2_Click()
icarcalc2.Checked = Not (icarcalc2.Checked)
End Sub


Private Sub lablshp_Click(Index As Integer)
On Error GoTo handle
lblshape_Click Index

Exit Sub
handle:
Exit Sub
End Sub

Private Sub lblshape_Click(Index As Integer)
'fac un modulo 7, cu un static ?
On Error GoTo handle
Static itest As Double
itest = itest + 1
lblshape(Index).Caption = CStr(CInt(1 + (itest Mod 7)))

Select Case CInt((itest Mod 7))
Case 0, 1
lablshp(Index).Caption = "circle"
Case 2
lablshp(Index).Caption = "cross"
Case 3
lablshp(Index).Caption = "x"
Case 4, 5
lablshp(Index).Caption = "square"
Case 6
lablshp(Index).Caption = "diamond"
End Select




Exit Sub
handle:
itest = 1
Exit Sub
End Sub

Private Sub mnu_comp1_Click()
mnu_comp1.Checked = Not (mnu_comp1.Checked)
End Sub

Private Sub mnu_comp2_Click()
mnu_comp2.Checked = Not (mnu_comp2.Checked)
End Sub

Private Sub mnu_comp3_Click()
mnu_comp3.Checked = Not (mnu_comp3.Checked)
End Sub

Private Sub mnu_comp4_Click()
mnu_comp4.Checked = Not (mnu_comp4.Checked)
End Sub


Private Sub mnu_curve1_Click()
mnu_curve1.Checked = Not (mnu_curve1.Checked)
End Sub

Private Sub mnu_curve2_Click()
mnu_curve2.Checked = Not (mnu_curve2.Checked)

End Sub

Sub mnu_draw_Click()
graphics.Cls
redeseneaza
End Sub

Private Sub mnu_exit_Click()
gindicator(2) = False
Unload parameters
Unload Me
End Sub

Private Sub mnu_export_Click()
Dim i As Integer, j As Integer
On Error GoTo handleit:
Screen.MousePointer = 11
Dim xg(200) As Double, yg(200) As Double, eroare As Boolean
Dim numarfisier As Integer, itest As Integer
numarfisier = FreeFile
Call deschide_fisier(CurDir$ + "\", 2, " text file (*.txt) |*.txt| results file (*.out) |*.out| show all (*.*) |*.*", 2)
If no_output Then Exit Sub
Open outputfile For Output Access Write As #numarfisier
xstart = CDbl(graphics.txtr(0).Text)
ystart = CDbl(graphics.txtr(2).Text)
xend = CDbl(graphics.txtr(1).Text)
yend = CDbl(graphics.txtr(3).Text)
If data_editor.tabdata.Caption = "CRTA" Then
Print #numarfisier, "CRTA data - curve 1 (alpha, temp./K)"
If mnu_curve1.Checked Then
For i = 1 To ipoints1
Print #numarfisier, Val(dalpha(i)), Val(dtempk(i))
Next i
End If
If mnu_curve2.Checked Then
Print #numarfisier, "CRTA data - curve 1 (alpha, temp./K)"
For i = ipoints1 + 2 To ipoints1 + ipoints2 + 1
Print #numarfisier, Val(dalpha(i)), Val(dtempk(i))
Next i
End If
If icarcalc1.Checked Then
j% = 1
Print #numarfisier, "CRTA data - interpolated data for curve 1 (alpha, temp./K)"
For i = 1 To nrint
Print #numarfisier, Val(dalpha(1) + (i - 1) * (dalpha(ipoints1) - dalpha(1)) / nrint), Val(igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i)))
Next i
End If
If icarcalc2.Checked Then
j% = 2
Print #numarfisier, "CRTA data - interpolated data for curve 2 (alpha, temp./K)"
For i = 1 To nrint
Print #numarfisier, Val(dalpha(1) + (i - 1) * (dalpha(ipoints1) - dalpha(1)) / nrint), Val(igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i)))
Next i
End If

If mnu_simulated.Checked Then
Print #numarfisier, "CRTA data - simulated data(alpha, temp./K)"
    For i = 1 To 200
    Print #numarfisier, Val(xgraf(i, 2)), Val(xgraf(i, 2))
    Next i
End If
Else
    If mnu_view_comp.Enabled Then
   
        If mnu_comp1.Checked Then
        Print #numarfisier, "Computed data - " & CStr(mnu_comp1.Caption)
        For i = 1 To 200
        Print #numarfisier, Val(xgraf(i, 3)), Val(ygraf(i, 3))
        Next i
        End If
    
        If mnu_comp2.Checked Then
        Print #numarfisier, "Computed data - " & CStr(mnu_comp2.Caption)
        For i = 1 To 200
        Print #numarfisier, Val(xgraf(i, 4)), Val(ygraf(i, 4))
        Next i
        End If
        
        If mnu_comp3.Checked Then
        Print #numarfisier, "Computed data - " & CStr(mnu_comp4.Caption)
        For i = 1 To 200
        Print #numarfisier, Val(xgraf(i, 5)), Val(ygraf(i, 5))
        Next i
        End If
        
        If mnu_comp4.Checked Then
        Print #numarfisier, "Computed data - " & CStr(mnu_comp4.Caption)
        For i = 1 To 200
        Print #numarfisier, Val(xgraf(i, 6)), Val(ygraf(i, 6))
        Next i
        End If
   End If

    If mnu_view_exp.Checked Then
        Print #numarfisier, "Experimental data (temp./K, alpha) "
        For i = 1 To ipoints
        Print #numarfisier, Val(dtempk(i)), Val(dalpha(i))
        Next i
    End If

    If mnu_simulated.Checked Then
        Print #numarfisier, "Simulated data (temp./K, alpha) "
        For i = 1 To 200
        Print #numarfisier, Val(xgraf(i, 2)), Val(ygraf(i, 2))
        Next i
      End If
End If
DoEvents
graphics.MousePointer = 0
Close
Exit Sub
handleit:
If Not (Err.Number = 32755) Then MsgBox "Unexpected error." & vbCrLf & CStr(Err.Description)
Close
graphics.MousePointer = 0
Exit Sub
End Sub

Private Sub mnu_follow_Click()
If (mnu_follow.Checked = False) Then
graphics.MousePointer = 2
'xy(0).Left = graphics.Left
'xy(1).Left = graphics.Left
xy(0).Visible = True
xy(1).Visible = True
Else
graphics.MousePointer = 0
xy(0).Visible = False
xy(1).Visible = False
End If
mnu_follow.Checked = Not (mnu_follow.Checked)
DoEvents
End Sub

Private Sub mnu_graph_copy_Click()
Clipboard.SetData graphics.Image
End Sub


Private Sub mnu_help_Click()
retval = WinHelp(graphics.hwnd, "versat10.hlp", HELP_KEY, CLng(0))

End Sub

Private Sub mnu_picture_Click()
On Error GoTo localhandle
Dim numarfisier As Integer
Dim itest As Integer
Dim dtest As Single
numarfisier = FreeFile
Call deschide_fisier(CurDir$ + "\", 2, " bmp file (*.bmp) |*.bmp| show all (*.*) |*.*", 1)
'trebuie testat raspunsul la cancel si alte erori, asta o fac prin no_otput
If no_output Then Exit Sub
Open outputfile For Output Access Write As #numarfisier
SavePicture graphics.Image, outputfile
Exit Sub
localhandle:
Exit Sub
End Sub

Private Sub mnu_sim_Click()
lbl(1).Enabled = True: txtp(1).Enabled = True: txtl(1).Enabled = True: coll(1).Enabled = True: colp(1).Enabled = True
lblshape(1).Visible = True
apelsimulare = True
mnu_simulated.Enabled = True
mnu_simulated.Checked = True
gindicator(2) = True
parameters.Show
End Sub

Private Sub mnu_simulated_Click()
mnu_simulated.Checked = Not (mnu_simulated.Checked)
End Sub

Private Sub mnu_style_Click()
If mnu_follow.Checked Then mnu_follow_Click
mnu_graph.Enabled = False
mnu_edit.Enabled = False
mnu_view.Enabled = False
graphics.BackColor = &HC0C0C0
graphics.Cls
graphics.BackColor = &HC0C0C0
If gindicator(2) Then
i = 1
lbl(3 + i).Visible = gindicator(i + 1): txtp(i).Visible = gindicator(i + 1): txtl(i).Visible = gindicator(i + 1): coll(i).Visible = gindicator(i + 1): colp(i).Visible = gindicator(i + 1)
End If

Style.Visible = True
For j% = 0 To 6
lblshape(j%).Visible = txtp(j%).Visible
lablshp(j%).Visible = lblshape(j%).Visible
Next j%
txtr(0).Text = CStr(xstart)
txtr(1).Text = CStr(xend)
txtr(2).Text = CStr(ystart)
txtr(3).Text = CStr(yend)

End Sub

Private Sub mnu_view_diff_Click()
'trebuie sa afisez formul de simulare pentru introducerea datelor
'daca in formul de simulare se apasa cancel desabled difference (trebuie sa am
'oricum datele experimentale)
'sterg pe toate celelalte checked
If mnu_view_diff.Checked Then
'il sterg pe checked si fac enabled celelalte
mnu_view_diff.Checked = False
Else
'il pun pe checked aici si le fac pe toate celelalte disabled
mnu_view_diff.Checked = True
End If
End Sub
Sub mnu_view_exp_Click()
mnu_view_exp.Checked = Not (mnu_view_exp.Checked)
'If mnu_view_comp.Checked = False Then mnu_view_diff.Checked = False
End Sub

Private Sub print_Click()
On Error GoTo handleit
Me.MousePointer = 11
Dim xg(200) As Double, yg(200) As Double, eroare As Boolean
If mnu_follow.Checked Then mnu_follow_Click
main_display.comdialog1.Flags = cdlPDReturnDC Or cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoWarning Or cdlPDUseDevModeCopies Or cdlPDAllPages Or cdlPDNoSelection
main_display.comdialog1.ShowPrinter
'Printer.DrawWidth = 1
' Printer.Orientation = vbPRORLandscape
Printer.ScaleTop = 0
Printer.ScaleLeft = 0
unitx = Printer.Width
unity = Printer.Height
For k% = 1 To main_display.comdialog1.Copies
'graficul il trasez doar in coordonate de temp, alpha
'desenez un frame
Printer.DrawWidth = 2
DoEvents
Printer.Line (Printer.Width / 8, Printer.Height * 7 / 8)-(Printer.Width * 7 / 8, Printer.Height / 8), 0, B
If (graphics.chk(0).Value = 1) Then
For i% = 0 To 5
Printer.Line (Printer.Width / 8 + i% * (Printer.Width - Printer.Width / 4) / 5, Printer.Height / 8)-Step(0, Printer.Height - Printer.Height / 4), &H0
Next i%
End If

If (graphics.chk(1).Value = 1) Then
For i% = 0 To 4
Printer.Line (Printer.Width / 8, Printer.Height / 8 + i% * (Printer.Height - Printer.Height / 4) / 4)-Step(Printer.Width - Printer.Width / 4, 0), &H0
Next i%
End If

xstart = CDbl(graphics.txtr(0).Text)
ystart = CDbl(graphics.txtr(2).Text)
xend = CDbl(graphics.txtr(1).Text)
yend = CDbl(graphics.txtr(3).Text)

pas# = (xend - xstart) / 5#
For i% = 1 To 5
Printer.CurrentY = 7.25 / 8 * Printer.Height
Printer.CurrentX = -0.025 * Printer.Width + 1 / 8 * Printer.Width + (i% - 1) * (Printer.Width * 3 / 4) / 5
Printer.Print Format$((xstart) + (i% - 1) * pas#, "##0.00")
Next i%
Printer.CurrentY = 7.25 / 8 * Printer.Height
Printer.CurrentX = Printer.Width * 7 / 8
Printer.Print CStr(graphics.txt(0).Text)



pas# = (yend - ystart) / 4#
For i% = 1 To 4
Printer.CurrentX = Printer.Width * 0.5 / 8
Printer.CurrentY = -Printer.Height * 0.005 + Printer.Height * 7 / 8 - (i% - 1) * (Printer.Height * 3 / 4) / 4
Printer.Print Format$((ystart) + (i% - 1) * pas#, "##0.00")
Next i%
Printer.CurrentY = Printer.Height / 8
Printer.CurrentX = Printer.Width * 0.5 / 8
Printer.Print CStr(graphics.txt(1).Text)



If data_editor.tabdata.Caption = "CRTA" Then
    If mnu_curve1.Checked Then
        For i% = 1 To ipoints1: xg(i%) = dalpha(i%): yg(i%) = dtempk(i%): Next i%
        Call printdesen(ipoints1, xg(), yg(), xstart, xend, ystart, yend, pointsize(1), pointcolor(1), linesize(1), linecolor(1), CInt(lblshape(0).Caption), eroare)
    End If
    If mnu_curve2.Checked Then
        For i% = ipoints1 + 2 To ipoints1 + ipoints2 + 1: xg(i% - ipoints1 - 1) = dalpha(i%): yg(i% - ipoints1 - 1) = dtempk(i%): Next i%
        Call printdesen(ipoints2, xg(), yg(), xstart, xend, ystart, yend, pointsize(2), pointcolor(2), linesize(2), linecolor(2), CInt(lblshape(1).Caption), eroare)
    End If
    If icarcalc1.Checked Then
    j% = 1
    For i% = 1 To nrint: xg(i%) = dalpha(1) + (i% - 1) * (dalpha(ipoints1) - dalpha(1)) / nrint: yg(i%) = igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i%)): Next i%
    Call printdesen(nrint, xg(), yg(), xstart, xend, ystart, yend, pointsize(3), pointcolor(3), linesize(3), linecolor(3), CInt(lblshape(2).Caption), eroare)
    End If
    If icarcalc2.Checked Then
    j% = 2
        For i% = 1 To nrint: xg(i%) = dalpha(ipoints1 + 2) + (i% - 1) * (dalpha(ipoints1 + ipoints2 + 1) - dalpha(ipoints1 + 2)) / nrint: yg(i%) = igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i%)): Next i%
        Call printdesen(nrint, xg(), yg(), xstart, xend, ystart, yend, pointsize(4), pointcolor(4), linesize(4), linecolor(4), CInt(lblshape(3).Caption), eroare)
    End If
    If mnu_simulated.Checked Then
        For i% = 1 To 200: xg(i%) = xgraf(i%, 2): yg(i%) = ygraf(i%, 2): Next i%
        Call printdesen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(7), pointcolor(7), linesize(7), linecolor(7), CInt(lblshape(6).Caption), eroare)
    End If
    If icar_dif.Checked Then
        j% = 1
        For i% = 1 To ipoints1: xg(i%) = dalpha(i%): yg(i%) = dtempk(i%) - igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i%)): Next i%
        Call printdesen(ipoints1, xg(), yg(), xstart, xend, ystart, yend, pointsize(5), pointcolor(5), linesize(5), linecolor(5), CInt(lblshape(4).Caption), eroare)
        j% = 2
        For i% = ipoints1 + 2 To ipoints1 + ipoints2 + 1: xg(i%) = dalpha(i% - ipoints1 - 1): yg(i%) = dtempk(i% - ipoints1 - 1) - igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), xg(i%)): Next i%
        Call printdesen(ipoints2, xg(), yg(), xstart, xend, ystart, yend, pointsize(6), pointcolor(6), linesize(6), linecolor(6), CInt(lblshape(5).Caption), eroare)
        End If
Else
    If mnu_view_comp.Enabled Then
    
        If mnu_comp1.Checked Then
        For i% = 1 To 200
        xg(i%) = xgraf(i%, 3): yg(i%) = ygraf(i%, 3)
        Next i%
        Call printdesen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(3), pointcolor(3), linesize(3), linecolor(3), CInt(lblshape(2).Caption), eroare)
        End If
        If mnu_comp2.Checked Then
        For i% = 1 To 200
        xg(i%) = xgraf(i%, 4): yg(i%) = ygraf(i%, 4)
        Next i%
        Call printdesen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(4), pointcolor(4), linesize(4), linecolor(4), CInt(lblshape(3).Caption), eroare)
        End If
        
        If mnu_comp3.Checked Then
        For i% = 1 To 200
        xg(i%) = xgraf(i%, 5): yg(i%) = ygraf(i%, 5)
        Next i%
        Call printdesen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(5), pointcolor(5), linesize(5), linecolor(5), CInt(lblshape(4).Caption), eroare)
        End If
        
        If mnu_comp4.Checked Then
        For i% = 1 To 200
        xg(i%) = xgraf(i%, 6): yg(i%) = ygraf(i%, 6)
        Next i%
        Call printdesen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(6), pointcolor(6), linesize(6), linecolor(6), CInt(lblshape(5).Caption), eroare)
        End If
    
    End If

    If mnu_view_exp.Checked Then
    Call printdesen(ipoints, dtempk(), dalpha(), xstart, xend, ystart, yend, pointsize(1), pointcolor(1), linesize(1), linecolor(1), CInt(lblshape(0).Caption), eroare)
    End If

    If mnu_simulated.Checked Then
    For i% = 1 To 200
    xg(i%) = xgraf(i%, 2): yg(i%) = ygraf(i%, 2)
    Next i%
    Call printdesen(200, xg(), yg(), xstart, xend, ystart, yend, pointsize(2), pointcolor(2), linesize(2), linecolor(2), CInt(lblshape(1).Caption), eroare)
    End If

End If
Next k%
Me.MousePointer = 0
Printer.EndDoc
Exit Sub
handleit:
Me.MousePointer = 0

Exit Sub


End Sub

