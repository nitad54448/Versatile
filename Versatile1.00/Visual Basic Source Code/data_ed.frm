VERSION 4.00
Begin VB.Form data_editor 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Versatil - data editor"
   ClientHeight    =   5505
   ClientLeft      =   1470
   ClientTop       =   1335
   ClientWidth     =   8370
   ForeColor       =   &H00000000&
   Height          =   6195
   HelpContextID   =   2000
   Left            =   1410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Top             =   705
   Width           =   8490
   Begin VB.TextBox txtstd 
      BackColor       =   &H00FFFFFF&
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
      Height          =   285
      HelpContextID   =   2040
      Left            =   7440
      MaxLength       =   5
      TabIndex        =   55
      Text            =   "1.0"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
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
      Height          =   325
      HelpContextID   =   2030
      Left            =   4920
      MaxLength       =   55
      MousePointer    =   3  'I-Beam
      TabIndex        =   53
      Text            =   "Experiment:"
      Top             =   5040
      Width           =   3255
   End
   Begin VB.ComboBox lst2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "data_ed.frx":0000
      Left            =   6600
      List            =   "data_ed.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   480
      Width           =   1325
   End
   Begin VB.ComboBox lst1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "data_ed.frx":0004
      Left            =   5160
      List            =   "data_ed.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   480
      Width           =   1455
   End
   Begin TabDlg.SSTab tabdata 
      Height          =   5295
      HelpContextID   =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8065
      _ExtentY        =   9335
      _StockProps     =   15
      Caption         =   "Regression"
      TabsPerRow      =   4
      Tab             =   2
      TabOrientation  =   0
      Tabs            =   4
      Style           =   1
      TabMaxWidth     =   0
      TabHeight       =   617
      WordWrap        =   0   'False
      TabCaption(0)   =   "Integral"
      Tab(0).ControlCount=   15
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblint(1)"
      Tab(0).Control(1)=   "lblint(2)"
      Tab(0).Control(2)=   "lblint(3)"
      Tab(0).Control(3)=   "lblint(0)"
      Tab(0).Control(4)=   "lblinteg"
      Tab(0).Control(5)=   "Line7(4)"
      Tab(0).Control(6)=   "Line7(6)"
      Tab(0).Control(7)=   "txtint(1)"
      Tab(0).Control(8)=   "txtint(2)"
      Tab(0).Control(9)=   "txtint(3)"
      Tab(0).Control(10)=   "chkint(0)"
      Tab(0).Control(11)=   "chkint(1)"
      Tab(0).Control(12)=   "chkint(2)"
      Tab(0).Control(13)=   "chkint(3)"
      Tab(0).Control(14)=   "txtint(0)"
      TabCaption(1)   =   "Differential"
      Tab(1).ControlCount=   15
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbldif(0)"
      Tab(1).Control(1)=   "lbldif(1)"
      Tab(1).Control(2)=   "lbldif(2)"
      Tab(1).Control(3)=   "lbldif(3)"
      Tab(1).Control(4)=   "lbldifer"
      Tab(1).Control(5)=   "Line7(5)"
      Tab(1).Control(6)=   "Line7(7)"
      Tab(1).Control(7)=   "chkdif(0)"
      Tab(1).Control(8)=   "chkdif(1)"
      Tab(1).Control(9)=   "chkdif(2)"
      Tab(1).Control(10)=   "chkdif(3)"
      Tab(1).Control(11)=   "txtdif(0)"
      Tab(1).Control(12)=   "txtdif(1)"
      Tab(1).Control(13)=   "txtdif(2)"
      Tab(1).Control(14)=   "txtdif(3)"
      TabCaption(2)   =   "Regression"
      Tab(2).ControlCount=   10
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "lblconv"
      Tab(2).Control(2)=   "lblreg(2)"
      Tab(2).Control(3)=   "Line7(1)"
      Tab(2).Control(4)=   "Line7(9)"
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(6)=   "chkreg(0)"
      Tab(2).Control(7)=   "chkreg(1)"
      Tab(2).Control(8)=   "chkreg(2)"
      Tab(2).Control(9)=   "txtreg"
      TabCaption(3)   =   "CRTA"
      Tab(3).ControlCount=   21
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblcrta(0)"
      Tab(3).Control(1)=   "lblcrta(3)"
      Tab(3).Control(2)=   "lblcrta(4)"
      Tab(3).Control(3)=   "lblcrta(1)"
      Tab(3).Control(4)=   "lblcrta(2)"
      Tab(3).Control(5)=   "lblcrta(6)"
      Tab(3).Control(6)=   "lblcrta(5)"
      Tab(3).Control(7)=   "lblcrta(7)"
      Tab(3).Control(8)=   "lblcrta(8)"
      Tab(3).Control(9)=   "lblc"
      Tab(3).Control(10)=   "Line1(0)"
      Tab(3).Control(11)=   "Line1(1)"
      Tab(3).Control(12)=   "txtcrta(0)"
      Tab(3).Control(13)=   "txtcrta(1)"
      Tab(3).Control(14)=   "txtcrta(2)"
      Tab(3).Control(15)=   "txtcrta(3)"
      Tab(3).Control(16)=   "txtcrta(4)"
      Tab(3).Control(17)=   "txtcrta(5)"
      Tab(3).Control(18)=   "txtcrta(6)"
      Tab(3).Control(19)=   "txtcrta(7)"
      Tab(3).Control(20)=   "txtcrta(8)"
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   8
         Left            =   -72840
         MaxLength       =   8
         TabIndex        =   39
         Text            =   "4"
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   7
         Left            =   -72840
         MaxLength       =   5
         TabIndex        =   38
         Text            =   "50"
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtreg 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   51
         Text            =   "5.0"
         Top             =   2760
         Width           =   735
      End
      Begin VB.CheckBox chkreg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[-ln(1-alpha)]^p"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   43
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkreg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "(1-alpha)^n"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   42
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkreg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "alpha^m"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   41
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   6
         Left            =   -72840
         MaxLength       =   8
         TabIndex        =   37
         Text            =   "0.02"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   5
         Left            =   -72840
         MaxLength       =   5
         TabIndex        =   36
         Text            =   "1.0"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   4
         Left            =   -72840
         MaxLength       =   5
         TabIndex        =   35
         Text            =   "0.0"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   3
         Left            =   -72840
         MaxLength       =   6
         TabIndex        =   34
         Text            =   "0.05"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   2
         Left            =   -72840
         MaxLength       =   5
         TabIndex        =   33
         Text            =   "1.0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   288
         Index           =   1
         Left            =   -72840
         MaxLength       =   6
         TabIndex        =   32
         Text            =   "0.0"
         Top             =   1560
         Width           =   732
      End
      Begin VB.TextBox txtcrta 
         BackColor       =   &H00FFFFFF&
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
         Height          =   288
         Index           =   0
         Left            =   -72840
         MaxLength       =   6
         TabIndex        =   31
         Text            =   "0.01"
         Top             =   1200
         Width           =   732
      End
      Begin VB.TextBox txtdif 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   5
         TabIndex        =   21
         Text            =   "5.0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtdif 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   5
         TabIndex        =   20
         Text            =   "0.1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtdif 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   5
         TabIndex        =   19
         Text            =   "3.00"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtdif 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   5
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox chkdif 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fatu (DTA)"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   -72240
         TabIndex        =   17
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox chkdif 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Piloyan"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   -74520
         TabIndex        =   16
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CheckBox chkdif 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Freeman-Carroll"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   -74520
         TabIndex        =   15
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CheckBox chkdif 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Achar "
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   -74520
         TabIndex        =   14
         Top             =   3360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtint 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox chkint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Urbanovici-Segal"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   -74520
         TabIndex        =   12
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CheckBox chkint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "van Krevelen"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   -74520
         TabIndex        =   11
         Top             =   4080
         Width           =   1695
      End
      Begin VB.CheckBox chkint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Flynn-Wall"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   -74520
         TabIndex        =   10
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CheckBox chkint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Coats-Redfern"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   -74520
         TabIndex        =   9
         Top             =   3360
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtint 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "5.0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtint 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "0.1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtint 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "3.00"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   -74760
         X2              =   -70680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   -74760
         X2              =   -70680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblc 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ICAR 2"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   -74640
         TabIndex        =   59
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   " ( pseudo-Inverse Matrix method )"
         ForeColor       =   &H00000000&
         Height          =   372
         Left            =   480
         TabIndex        =   58
         Top             =   3480
         Width           =   3492
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Polynom degree"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   8
         Left            =   -74760
         TabIndex        =   57
         Top             =   4320
         Width           =   1692
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Intervals"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   7
         Left            =   -74640
         TabIndex        =   56
         Top             =   3960
         Width           =   1572
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   240
         X2              =   4320
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   -74760
         X2              =   -70680
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   -74760
         X2              =   -70680
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   5
         X1              =   -74760
         X2              =   -70680
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   4
         X1              =   -74760
         X2              =   -70680
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   240
         X2              =   4320
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "mass for alpha=1"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   5
         Left            =   -74880
         TabIndex        =   52
         Top             =   3120
         Width           =   1812
      End
      Begin VB.Label lblreg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Heating rate K/min"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   50
         Top             =   2760
         Width           =   1692
      End
      Begin VB.Label lbldifer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obtain DTG from TG data"
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
         Height          =   375
         Left            =   -74640
         TabIndex        =   49
         Top             =   2520
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblinteg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obtain TG by integration"
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
         Height          =   375
         Left            =   -74640
         TabIndex        =   48
         Top             =   2520
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Excluded domain"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   6
         Left            =   -74880
         TabIndex        =   40
         Top             =   3600
         Width           =   1812
      End
      Begin VB.Label lblconv 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[(1-alpha)^n]"
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
         Left            =   1200
         TabIndex        =   45
         Top             =   600
         Width           =   3252
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conversion function"
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
         Height          =   492
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "mass for alpha=1"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   2
         Left            =   -74880
         TabIndex        =   30
         Top             =   1920
         Width           =   1812
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "mass for alpha=0"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   1
         Left            =   -74880
         TabIndex        =   29
         Top             =   1560
         Width           =   1812
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "mass for alpha=0"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   4
         Left            =   -74880
         TabIndex        =   28
         Top             =   2760
         Width           =   1812
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "2nd rate (mg/min)"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   3
         Left            =   -74880
         TabIndex        =   27
         Top             =   2400
         Width           =   1812
      End
      Begin VB.Label lblcrta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "1st rate (mg/min)"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   -74880
         TabIndex        =   26
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label lbldif 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rate (K/min)"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   25
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lbldif 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Step"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lbldif 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Final order"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lbldif 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Initial order"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblint 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Initial order"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblint 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rate (K/min)"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblint 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Step"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblint 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Final order"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Est. error, in alpha , %"
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
      Left            =   5040
      TabIndex        =   54
      Top             =   120
      Width           =   2295
   End
   Begin MSGrid.Grid Grid1 
      Height          =   4395
      HelpContextID   =   2020
      Left            =   4935
      TabIndex        =   13
      Top             =   480
      Width           =   3230
      _Version        =   65536
      _ExtentX        =   5689
      _ExtentY        =   7752
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         name            =   "MS Sans Serif"
         charset         =   238
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Rows            =   18
      Cols            =   3
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      HelpContextID   =   2101
      Begin VB.Menu mnu_cancel 
         Caption         =   "O&K"
      End
      Begin VB.Menu mnu_save 
         Caption         =   "&Save file"
      End
      Begin VB.Menu mnu_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_open 
         Caption         =   "&Open file"
      End
      Begin VB.Menu mnu_ascii 
         Caption         =   "&Import ASCII file"
      End
      Begin VB.Menu mnu_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_print 
         Caption         =   "&Print data"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_export 
         Caption         =   "&Export ASCII file"
      End
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "&Edit"
      HelpContextID   =   2201
      Begin VB.Menu mnu_import 
         Caption         =   "&Import picture"
      End
      Begin VB.Menu m_ 
         Caption         =   "-"
      End
      Begin VB.Menu canceledit 
         Caption         =   "&Cancel edit"
      End
      Begin VB.Menu mnu_discard 
         Caption         =   "&Discard data"
      End
      Begin VB.Menu ms 
         Caption         =   "-"
      End
      Begin VB.Menu smooth 
         Caption         =   "Smooth Y data"
      End
      Begin VB.Menu shiftx 
         Caption         =   "Shift X data"
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu swap 
         Caption         =   "&Swap columns"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "data_editor"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub canceledit_Click()
main_display!menu_comp_param.Enabled = False
main_display!menu_comp_param.Enabled = False
main_display!menu_print_data.Enabled = False
main_display!save_data.Enabled = False
gindicator(1) = False
data_editor.Hide
main_display.Show
End Sub

Private Sub chkdif_Click(Index As Integer)
On Error GoTo handleit
For i% = 0 To 2
If chkdif(i%).Value Then chkdif(3).Value = 0
Next i%
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub chkreg_Click(Index As Integer)
'il las pe bizon sa innegreasca tot ce poate, dau cu flitul dupa verif
On Error GoTo handleit
lblconv.Caption = ""
If chkreg(0).Value Then lblconv.Caption = lblconv.Caption + "[alpha^m]"
If chkreg(1).Value Then lblconv.Caption = lblconv.Caption + "[(1-alpha)^n]"
If chkreg(2).Value Then lblconv.Caption = lblconv.Caption + "[(-ln(1-alpha))^p]"
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub Form_Activate()
lst2_Click
End Sub

Private Sub Form_Load()
On Error GoTo handleit
Me.Caption = nume_prog & " - data editor"
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
'scriu in lst1 si lst2
lst1.AddItem "temp. /C"
lst1.AddItem "temp. /K"
lst1.ListIndex = 0
lst2.AddItem "TG values"
lst2.AddItem "alpha"
lst2.AddItem "DTG values"
lst2.AddItem "DTA values"
lst2.ListIndex = 3

Grid1.Rows = 20 'se incepe de la 0, atentie
Grid1.ColWidth(0) = 1.6 * Grid1.Width / 15.3
Grid1.RowHeight(0) = Grid1.Height / 16
Grid1.ColWidth(1) = 6 * Grid1.Width / 15.3
Grid1.ColWidth(2) = 6 * Grid1.Width / 15.3
lst1.Width = Grid1.ColWidth(1)
lst1.Left = Grid1.Left + Grid1.ColWidth(0)
lst2.Width = Grid1.ColWidth(2)

'aranjez in pagina
Grid1.FixedAlignment(0) = 2
Grid1.ColAlignment(1) = 2
Grid1.ColAlignment(2) = 2
For i% = 1 To 7
gindicator(i%) = False
Next i%
Grid1.Row = 1
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
mnu_cancel_Click
Cancel = 1 'daca cancel=0 atunci o descarc altfel nu
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
'maresc numarul de linii daca este cazul
On Error GoTo handleit
If Grid1.Row = 0 Then SendKeys "{down}"
If Grid1.Row > Grid1.Rows - 3 Then Grid1.Rows = Grid1.Rows + 1
Select Case KeyAscii
    Case 49 To 57 'de la 1 la 9
        Grid1.Text = Grid1.Text + Chr(KeyAscii)
    Case 8 'backspace
        If Len(Grid1.Text) > 0 Then Grid1.Text = Left$(Grid1.Text, Len(Grid1.Text) - 1)
    Case 13 'enterul asta cretin
        If Grid1.Row < Grid1.Rows - 2 Then
            If Grid1.Col = 1 Then SendKeys "{right}"
            If Grid1.Col = 2 Then Grid1.Col = 1: SendKeys "{down}"
        End If
    Case 48 '0
        If Grid1.Text = "0" Then Exit Sub
        Grid1.Text = Grid1.Text + "0"
    Case 46, 48 '. si ,
        If InStr(Grid1.Text, ".") Then Exit Sub
        Grid1.Text = Grid1.Text + "."
    Case 45 '-
    If Len(Grid1.Text) > 0 Then
    If Left$(Grid1.Text, 1) = "-" Then
        Grid1.Text = Right$(Grid1.Text, Len(Grid1.Text) - 1)
        Else
        Grid1.Text = "-" & Grid1.Text
    End If
    Else
    Grid1.Text = "-"
    End If
End Select
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
'shift=2 inseamna ctrl+
On Error GoTo handle
If Shift = 2 Then
Dim stest As String, itest As Integer
itest = Grid1.Row
Select Case KeyCode
    Case vbKeyInsert
'insereaza o linie goala
    For i = Grid1.Rows - 2 To Grid1.Row Step -1
    Grid1.Col = 1
    Grid1.Row = i: stest = Grid1.Text
    Grid1.Row = i + 1: Grid1.Text = stest
    Grid1.Col = 2
    Grid1.Row = i: stest = Grid1.Text
    Grid1.Row = i + 1: Grid1.Text = stest
    Next i
    Grid1.Row = itest
    For i = 1 To 2: Grid1.Col = i: Grid1.Text = "": Next i
    Case vbKeyDelete
'sterge o linie
    For i = Grid1.Row To Grid1.Rows - 2
    Grid1.Col = 1
    Grid1.Row = i + 1: stest = Grid1.Text
    Grid1.Row = i: Grid1.Text = stest
    Grid1.Col = 2
    Grid1.Row = i + 1: stest = Grid1.Text
    Grid1.Row = i: Grid1.Text = stest
    Next i
End Select
Grid1.Row = itest: Grid1.Col = 1
End If
DoEvents
Exit Sub
handle:
'eroare...
Exit Sub
End Sub
Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
If Button = vbRightButton Then
data_editor.PopupMenu mnu_file, , , , mnu_save
End If
Grid1.Row = Grid1.Row
End Sub

Private Sub lst2_Click()
    Select Case tabdata.Caption
    Case "Regression"
        Select Case lst2.Text
        Case "DTA values"
            MsgBox "DTA data are not accepted for this method." & vbCrLf & "Prepare your data as alpha, TG or DTG values."
            lst2.ListIndex = 0
        End Select
    Case "CRTA"
        Select Case lst2.Text
        Case "DTA values"
            MsgBox "DTA data are not accepted for this method." & vbCrLf & "Prepare your data as alpha or TG values."
            lst2.ListIndex = 0
        Case "DTG values"
            MsgBox "DTG data are not accepted for this method." & vbCrLf & "Prepare your data as alpha or TG values."
            lst2.ListIndex = 0
        End Select
Case "Integral", "Differential"

End Select


Select Case lst2.Text
Case "DTG values"
lblinteg.Visible = True
lbldifer.Visible = False
chkdif(0).Enabled = True
chkdif(1).Enabled = True
chkdif(2).Enabled = True
chkdif(3).Enabled = False

Case "DTA values"
lblinteg.Visible = True
lbldifer.Visible = False
chkdif(0).Enabled = False
chkdif(1).Enabled = False
chkdif(2).Enabled = False
chkdif(0).Value = False
chkdif(1).Value = False
chkdif(2).Value = False
chkdif(3).Enabled = True


Case "TG values"
lbldifer.Visible = True
lblinteg.Visible = False
chkdif(0).Enabled = True
chkdif(1).Enabled = True
chkdif(2).Enabled = True
chkdif(3).Enabled = False
chkdif(3).Value = False

Case Else
lblinteg.Visible = False
lbldifer.Visible = True
chkdif(0).Enabled = True
chkdif(1).Enabled = True
chkdif(2).Enabled = True
chkdif(3).Enabled = False
chkdif(3).Value = False
End Select
Exit Sub
End Sub

Private Sub mnu_about_Click()
apel = True
about.Show
End Sub

Private Sub mnu_ascii_Click()
On Error GoTo localhandle
'rutina ce introduce o serie de ascii
'prima linie nume_exp
Dim numarfisier As Integer, stest As String, itest As Integer, sa As Double, sb As Double
numarfisier = FreeFile
Call deschide_fisier(CurDir$ + "\", 1, " text file (*.txt) |*.txt| data file (*.dat) |*.dat| show all (*.*) |*.*", 1)
If no_input Then Exit Sub
Open inputfile For Input Access Read As #numarfisier

Line Input #numarfisier, stest
txtname.Text = Left$(stest, 50)
Grid1.Col = 1
itest = 0
'daca am ceva date le pun la sfarsit
Do
itest = itest + 1
Grid1.Row = itest
If Len(Grid1.Text) = 0 Then Exit Do
Loop Until itest = Grid1.Rows - 3
itest = itest - 1

Do While Not (EOF(numarfisier))
itest = itest + 1
Input #numarfisier, sa, sb
Grid1.Rows = Grid1.Rows + 1
Grid1.Row = itest
Grid1.Text = CStr(sa)
Grid1.Col = 2
Grid1.Text = CStr(sb)
Grid1.Col = 1
Loop
Close #numarfisier
Exit Sub
localhandle:
t = MsgBox("There may be an error...", vbOKOnly, nume_prog)
Close #numarfisier
Exit Sub
End Sub

Private Sub mnu_cancel_Click()
'fac verificarea datelor introduse,desi fac o verificare la load
'utilizatorul poate schimba ceva
'daca datele nu sunt conforme cu cerintele fac inactiv menu-ul de calcul
Dim eroare As Boolean
Call verif_data(eroare)
main_display!menu_comp_param.Enabled = True
main_display!menu_print_data.Enabled = True
main_display!save_data.Enabled = True
If eroare Then
i% = MsgBox("Quit the Data Editor window ?", vbQuestion + vbYesNo, nume_prog)
If i% = vbNo Then Exit Sub
main_display!menu_comp_param.Enabled = False
main_display!menu_comp_param.Enabled = False
main_display!menu_print_data.Enabled = False
main_display!save_data.Enabled = False
gindicator(1) = False
End If
data_editor.Hide
main_display.Show
If Not (eroare) Then gindicator(1) = True
End Sub

Private Sub mnu_discard_Click()
On Error GoTo handleit
itest% = MsgBox(" Discard and loss all the data ?", vbYesNo + vbExclamation + vbDefaultButton2, nume_prog)
If itest% = vbNo Then Exit Sub
'scade toate datele din memorie
curat_grid
Grid1.Row = 1
Grid1.Rows = 21
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub mnu_export_Click()
On Error GoTo localhandle
Dim numarfisier As Integer, itest As Integer, dtest As Double, i As Integer
numarfisier = FreeFile
Call deschide_fisier(CurDir$ + "\", 2, " text file (*.txt) |*.txt| data file (*.dat) |*.dat| show all (*.*) |*.*", 2)
'trebuie testat raspunsul la cancel si alte erori, asta o fac prin no_otput
If no_output Then Exit Sub
Open outputfile For Output Access Write As #numarfisier
Print #numarfisier, Left$(txtname.Text, 50)
itest = 0
For i = 1 To Grid1.Rows - 1
Grid1.Col = 1
itest = itest + 1
Grid1.Row = itest
dtest = Val(Grid1.Text)
Grid1.Col = 2
Print #numarfisier, Val(dtest), Val(Grid1.Text)
Next i
Close #numarfisier
Exit Sub
localhandle:
'a fost o eroare
t = MsgBox("There may be an error...", vbOKOnly, nume_prog)
Close
Exit Sub
End Sub

Private Sub mnu_help_Click()
retval = WinHelp(data_editor.hwnd, "versat10.hlp", HELP_KEY, CLng(0))
End Sub

Private Sub mnu_import_Click()
'la import picture
import_form.Show
import_form.mnu_open_Click
End Sub

Sub mnu_open_Click()
On Error GoTo handleit:
'arat data_editor pentru setfocus
'fac verificarile similar ca in verif
Me.MousePointer = 11
Dim numarfisier As Integer, stest As String, dtest(4) As Double, itest As Integer, i As Integer
numarfisier = FreeFile
Call deschide_fisier(CurDir$ + "\", 1, " text file (*.txt) |*.txt| data file (*.dat) |*.dat| show all (*.*) |*.*", 2)
If no_input Then Me.MousePointer = 0: Exit Sub
curat_grid
Open inputfile For Input Access Read As #numarfisier
'verifica daca este un fisier de date si introdu-l in data_ed
Line Input #numarfisier, stest
If Not (stest = "This is a data file for " & nume_prog) Then MsgBox "Error in the first line. You may encounter problems."
'daca este un fisier de date corect, citeste valorile si le introduce in data_editor
Line Input #numarfisier, stest
'acest stest trebuie sa aiba stringul Created in el altfel il flituiesc
If Not (Left$(stest, 7) = "Created") Then MsgBox "The second line is not correct (creation data). "
Line Input #numarfisier, stest
'stest are acum numele experimentului, primele 10 caractere sunt "Experiment"
If Not (Left$(stest, 10) = "Experiment") Then MsgBox "The third line is not correct, the experiment name." & vbCrLf & "Check the data file or edit a valid data file with" & nume_prog
txtname.Text = Left$(Right$(stest, Len(stest) - 17), 40)
'ar trebui sa fie un separator de o linie punctata, daca nu e
Line Input #numarfisier, stest
If Not (stest = linie) Then MsgBox "The data file may be incorrect (separation line)." & vbCrLf & "Check the data file or edit a valid data file with " & nume_prog
'incep sa citesc datele, integer etc
Line Input #numarfisier, stest
Select Case LCase$(Left$(stest, 4))
    Case "inte"
'citesc valorile care arata datele introduse, corespund cu pozitia in liste
    tabdata.SetFocus
    For i = 1 To 3
    If tabdata.Caption = "Integral" Then Exit For
    SendKeys "^{TAB}", True
    Next i
    
    Line Input #numarfisier, stest
    lst1.ListIndex = CInt(Left$(stest, 1)) - 1
    lst2.ListIndex = CInt(Mid$(stest, 2)) - 1
    'citesc ordinul initial, final, viteza
    Input #numarfisier, dtest(1), dtest(2), dtest(3), dtest(4)
    'nu fac acum verificarea valorilor lor
    For i = 1 To 4
    txtint(i - 1).Text = CStr(Val(dtest(i)))
    Next i
'citesc metodele alese, suma lor trebuie sa fie >0, <=3
    Line Input #numarfisier, stest
    If Len(stest) > 4 Then Err.Raise 1101, , " Incorrect data file. (in the integral methods descriptor)"
    For i = 1 To 4 'am doar 4 metode
        Select Case Val(Mid(stest, i, 1))
        Case 0, 1
        chkint(i - 1).Value = Val(Mid(stest, i, 1))
        Case Else
        Err.Raise 1101, , "Incorrect data file. (in the integral methods descriptor)"
        End Select
    Next i
'citesc sigy
    Input #numarfisier, stest
    sigy = Val(stest)
    If sigy < 0.05 Or sigy > 10 Then Err.Raise 1101, , "Incorrect value for estimated deviation, %"
    txtstd.Text = Format$(stest, "#0.00")
'trebuie sa citesc datele acum si sa le pun in grid
    itest = 0
    Do Until (EOF(numarfisier))
    Grid1.Col = 1
    itest = itest + 1
    Grid1.Row = itest
    Grid1.Rows = Grid1.Rows + 1
    Input #numarfisier, dtest(1), dtest(2)
    If dtest(1) < 0.1 Then Exit Do
    Grid1.Text = Val(dtest(1))
    Grid1.Col = 2
    Grid1.Text = Val(dtest(2))
    Loop

Case "diff"
    tabdata.SetFocus
    For i = 1 To 3
    If tabdata.Caption = "Differential" Then Exit For
    SendKeys "^{TAB}", True
    Next i
    
    Line Input #numarfisier, stest
    lst1.ListIndex = CInt(Left$(stest, 1)) - 1
    lst2.ListIndex = CInt(Mid$(stest, 2)) - 1
'citesc ordinul initial, final, viteza
    Input #numarfisier, dtest(1), dtest(2), dtest(3), dtest(4)
'nu fac acum verificarea valorilor lor
    If dtest(4) = 1 Then
        For i = 1 To 3
        dtest(i) = 0
        Next i
    End If
    For i = 1 To 4
    txtdif(i - 1).Text = CStr(Val(dtest(i)))
    Next i
'citesc metodele alese, suma lor trebuie sa fie >0, <=3
    Line Input #numarfisier, stest
    If Len(stest) > 4 Then MsgBox "There are problems in the methods descriptor"
    For i = 1 To 4 'am doar 4 metode
        Select Case Val(Mid(stest, i, 1))
        Case 0, 1
        chkdif(i - 1).Value = Val(Mid(stest, i, 1))
        Case Else
        Err.Raise 1101, , "Incorrect data file. (in the methods descriptor)"
        End Select
    Next i
'citesc sigy
    Input #numarfisier, stest
    sigy = Val(stest)
    If sigy < 0.05 Or sigy > 10 Then Err.Raise 1101, , "Incorrect value for estimated deviation, %"
    txtstd.Text = Format$(stest, "#0.00")
'trebuie sa citesc datele acum si sa le pun in grid
    itest = 0
    Do Until (EOF(numarfisier))
        Grid1.Col = 1
        itest = itest + 1
        Grid1.Row = itest
        Input #numarfisier, dtest(1), dtest(2)
        If dtest(1) < 0.1 Then Exit Do
        Grid1.Text = Val(dtest(1))
        Grid1.Col = 2
        Grid1.Text = Val(dtest(2))
    Loop

Case "regr"
    tabdata.SetFocus
    For i = 1 To 3
    If tabdata.Caption = "Regression" Then Exit For
    SendKeys "^{TAB}", True
    Next i
   
   Line Input #numarfisier, stest
    lst1.ListIndex = CInt(Left$(stest, 1)) - 1
    lst2.ListIndex = CInt(Mid$(stest, 2)) - 1
    'citesc  viteza
    Input #numarfisier, dtest(1)
    txtreg.Text = CStr(Val(dtest(1)))
'nu citesc metodele alese, totdeauna regresie, trebuie sa fie 1
    Line Input #numarfisier, stest
    If Len(stest) > 3 Then Err.Raise 1101, , " Incorrect data file. (in the conversion function descriptor.)"
For i = 1 To 3
chkreg(i - 1).Value = Val(Mid$(stest, i, 1))
Next i
'citesc sigy, desi nu il folosesc aici
    Input #numarfisier, stest
    sigy = Val(stest)
    If sigy < 0.05 Or sigy > 10 Then Err.Raise 1101, , "Incorrect value for estimated deviation, %"
    txtstd.Text = Format$(stest, "#0.00")
'trebuie sa citesc datele acum si sa le pun in grid
    itest = 0
    Do Until (EOF(numarfisier))
    Grid1.Col = 1
    itest = itest + 1
    Grid1.Row = itest
    Grid1.Rows = Grid1.Rows + 1
    Input #numarfisier, dtest(1), dtest(2)
    If dtest(1) < 0.1 Then Exit Do
    Grid1.Text = Val(dtest(1))
    Grid1.Col = 2
    Grid1.Text = Val(dtest(2))
    Loop

Case "crta"

    tabdata.SetFocus
    For i = 1 To 3
    If tabdata.Caption = "CRTA" Then Exit For
    SendKeys "^{TAB}", True
    Next i

    Line Input #numarfisier, stest
    lst1.ListIndex = CInt(Left$(stest, 1)) - 1
    lst2.ListIndex = CInt(Mid$(stest, 2)) - 1

'citesc sigy desi nu imi trebuie
 Line Input #numarfisier, stest
    sigy = CDbl(stest)
    If sigy < 0.05 Or sigy > 10 Then stest = "1.0"
    txtstd.Text = Format$(stest, "#0.00")
 
Dim vv(9) As Double
'icar  2
Input #numarfisier, vv(1), vv(2), vv(3), vv(4), vv(5), vv(6), vv(7), vv(8), vv(9)
For i = 1 To 9
txtcrta(i - 1).Text = CStr(vv(i))
Next i
'icar2
    itest = 0
    Do Until (EOF(numarfisier))
    Grid1.Rows = Grid1.Rows + 1
        Grid1.Col = 1
        itest = itest + 1
        Grid1.Row = itest
        Input #numarfisier, dtest(1)
        If dtest(1) < 0.1 Then Exit Do
        Grid1.Text = Val(dtest(1))
         Input #numarfisier, dtest(2)
        Grid1.Col = 2
        Grid1.Text = Val(dtest(2))
    
    Loop
   ' Input #numarfisier, stest 'string simplu
itest1% = itest - 1
    Do Until (EOF(numarfisier))
    Grid1.Rows = Grid1.Rows + 1
        Grid1.Col = 1
        itest = itest + 1
        Grid1.Row = itest
        Input #numarfisier, dtest(1), dtest(2)
        If dtest(1) < 0.1 Then Exit Do
        Grid1.Text = Val(dtest(1))
        Grid1.Col = 2
        Grid1.Text = Val(dtest(2))
    Loop
    If ((itest1% < 6) Or (itest - itest1%) < 6) Then Err.Raise 1101, , "You need two data sets, each one with more than 7 pairs of data (separe the sets by a space in the file, or in grid)."
    

Case Else
Err.Raise 1101, , "The data file is incorrect (the data type string is missing)." & vbCrLf & "Check the data file or edit a valid data file with" & nume_prog
End Select
Close #numarfisier
Me.MousePointer = 0
Exit Sub
handleit:
Me.MousePointer = 0
Close #numarfisier
If Err.Number = 62 Then Exit Sub
'If Not (Err.Number = 1101) Then
MsgBox "Error number " & CStr(Err.Number) & vbCrLf & CStr(Err.Description) & vbCrLf & "Unexpected error. Check the data or report the conditions to the author."
'Else
'i = MsgBox(CStr(Err.Description) & vbcrlf & "Check your data or report the conditions to the author.", vbInformation, ion, nume_prog)
'End If
Exit Sub 'exit sub face el err.clear
End Sub

Sub mnu_print_Click()
Me.MousePointer = 11
On Error GoTo handleit
Dim eroare As Boolean, space As String, txt As String
Call verif_data(eroare)
txt = ""
If eroare Then txt = "Attention: this is not a valid data file."
main_display.comdialog1.Flags = cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection Or cdlPDNoWarning Or cdlPDUseDevModeCopies
main_display.comdialog1.ShowPrinter
space = "   "
For i% = 1 To main_display.comdialog1.Copies
'tiparesc atatea copii cate cere bizonul
Printer.Print ""
Printer.Print ""
Printer.Print space & txt
Printer.Print space & nume_prog & " ,data file."
Printer.Print licenta
Printer.Print space & Format$(Now, "dddd, mmm d yyyy") + ", " + Format$(Now, "hh:mm:ss")
Printer.Print ""
Printer.Print space & linie
Printer.Print space & "Experiment name: " & txtname.Text
Printer.Print space & "Methods :" & CStr(tabdata.Caption)
Printer.Print ""
Select Case tabdata.Caption
Case "Integral"
Printer.Print space & "Selected procedures:"
For itest% = 0 To 3
If chkint(itest%).Value = 1 Then Printer.Print "         -" & chkint(itest%).Caption
Next itest%
Printer.Print space & "Initial reaction order: " & txtint(0).Text
Printer.Print space & "Final reaction order: " & txtint(1).Text
Printer.Print space & "Reaction order step: " & txtint(2).Text
Printer.Print space & "Heating rate, K/min: " & txtint(3).Text
Printer.Print space & "Estimated error in conversion: " & Format$(sigy, "#0.00#") & " %"
Printer.Print space & "Data :  " & lst1.List(lst1.ListIndex) & ", "; lst2.List(lst2.ListIndex)
For itest% = 1 To ipoints
Grid1.Row = itest%
Grid1.Col = 1
Printer.Print space & Grid1.Text;
Grid1.Col = 2
Printer.Print ",  " & Grid1.Text
Next itest%

Case "Differential"
Printer.Print space & "Selected procedures:"
For itest% = 0 To 3
If chkdif(itest%).Value = 1 Then Printer.Print "         -" & chkdif(itest%).Caption
Next itest%
Printer.Print space & "Initial reaction order: " & txtdif(0).Text
Printer.Print space & "Final reaction order: " & txtdif(1).Text
Printer.Print space & "Reaction order step: " & txtdif(2).Text
Printer.Print space & "Heating rate, K/min: " & txtdif(3).Text
Printer.Print space & "Estimated error in conversion: " & Format$(sigy, "#0.00#") & " %"
Printer.Print space & "Data :  " & lst1.List(lst1.ListIndex) & ", "; lst2.List(lst2.ListIndex)
For itest% = 1 To ipoints
Grid1.Row = itest%
Grid1.Col = 1
Printer.Print space & Grid1.Text;
Grid1.Col = 2
Printer.Print ", " & Grid1.Text
Next itest%


Case "Regression"
Printer.Print space & "Pseudo-Iverse matrix method."
Printer.Print space & "Conversion function : " & space & lblconv.Caption
Printer.Print space & "Heating rate, K/min: " & txtdif(3).Text
Printer.Print space & "Estimated error in conversion (not used in this method): " & Format$(sigy, "#0.00#") & " %"
Printer.Print space & "Data :  " & lst1.List(lst1.ListIndex) & ", "; lst2.List(lst2.ListIndex)
For itest% = 1 To ipoints
Grid1.Row = itest%
Grid1.Col = 1
Printer.Print space & Grid1.Text;
Grid1.Col = 2
Printer.Print ", " & Grid1.Text
Next itest%

Case "CRTA"
Printer.Print space & "ICAR 2 - procedure:"
Printer.Print space & " First data set:"
Printer.Print space & " Number of points: " & CStr(ipoints1)
Printer.Print space & " Decomposition rate (mg/min) : " & CStr(rate1)
Printer.Print space & " Decomposition rate (1/sec) : " & CStr(rate1)
Printer.Print space & " TG value for alpha=0 : " & CStr(ivm1)
Printer.Print space & " TG value for alpha=1 : " & CStr(ivm2)
Printer.Print space & "Data :  " & lst1.List(lst1.ListIndex) & ", "; lst2.List(lst2.ListIndex)
For itest% = 1 To ipoints1
Grid1.Row = itest%
Grid1.Col = 1
Printer.Print space & Grid1.Text;
Grid1.Col = 2
Printer.Print ", " & Grid1.Text
Next itest%

Printer.Print "  "
Printer.Print space & " Second data set:"
Printer.Print space & " Number of points: " & CStr(ipoints2)
Printer.Print space & " Decomposition rate (mg/min) : " & CStr(rate2)
Printer.Print space & " Decomposition rate (1/sec) : " & CStr(rate2)
Printer.Print space & " TG value for alpha=0 : " & CStr(iivm1)
Printer.Print space & " TG value for alpha=1 : " & CStr(iivm2)
Printer.Print "  "
Printer.Print space & "Data :  " & lst1.List(lst1.ListIndex) & ", "; lst2.List(lst2.ListIndex)
For itest% = ipoints1 + 2 To ipoints1 + ipoints2 + 1
Grid1.Row = itest%
Grid1.Col = 1
Printer.Print space & Grid1.Text;
Grid1.Col = 2
Printer.Print ", " & Grid1.Text
Next itest%
Printer.Print "  "
Printer.Print space & " Global parameters:"
Printer.Print space & " Number of intervals: " & CStr(nrint)
Printer.Print space & " Polynomial degree : " & CStr(grpol - 1)
Printer.Print space & " Excluded domain (alpha) : " & CStr(excl)
'
End Select
Printer.Print space & linie
Printer.Print ""
Next i%
Printer.EndDoc
Me.MousePointer = 0
Exit Sub
handleit:
Me.MousePointer = 0
If Err.Number = cdlCancel Then Exit Sub ' a apasat cancel
MsgBox CStr(Err.Description) & vbCrLf & "There were errors trying to print the data."
'Printer.EndDoc
'MsgBox "cancel"
'daca a apasat cancel iesi, daca nu da un mesaj de eroare inainte
Exit Sub
End Sub

Sub mnu_save_Click()
On Error GoTo handleit
'verific datele
Dim eroare As Boolean
Call verif_data(eroare)
If eroare Then Exit Sub
'scriu fisierul de date
Dim numarfisier As Integer
Dim itest As Integer
Dim dtest As Double
numarfisier = FreeFile
Call deschide_fisier(CurDir$ + "\", 2, " text file (*.txt) |*.txt| data file (*.dat) |*.dat| show all (*.*) |*.*", 2)
'trebuie testat raspunsul la cancel si alte erori, asta o fac prin no_otput
If no_output Then Exit Sub
Open outputfile For Output Access Write As #numarfisier
Print #numarfisier, "This is a data file for " & nume_prog
Print #numarfisier, "Created : " & Format$(Now, "dddd, mmm d yyyy") + ", " + Format$(Now, "h:m:s")
Print #numarfisier, "Experiment name: " & nume_exp
Print #numarfisier, linie
Select Case tabdata.Caption
Case "Integral"
Print #numarfisier, "integral methods"
Print #numarfisier, CStr(lst1.ListIndex + 1) & CStr(lst2.ListIndex + 1)
Print #numarfisier, Val(txtint(0).Text), txtint(1).Text, txtint(2).Text, txtint(3).Text
'indicator asupra metodelor alese, 1 pt on si 0 pt off, suma trebuie sa fie mai mare ca 0
Print #numarfisier, chkint(0).Value & chkint(1).Value & chkint(2).Value & chkint(3).Value '& chkint(4).Value
'tiparesc abaterea standard
Print #numarfisier, Val(sigy)
For itest = 1 To ipoints
Print #numarfisier, Val(dx(itest)) & " , " & spatiu & Val(dy(itest))
Next itest

Case "Differential"
Print #numarfisier, "differential methods"
Print #numarfisier, CStr(lst1.ListIndex + 1) & CStr(lst2.ListIndex + 1)
Print #numarfisier, Val(txtdif(0).Text), txtdif(1).Text, txtdif(2).Text, txtdif(3).Text
'indicator asupra metodelor alese, 1 pt on si 0 pt off, suma trebuie sa fie mai mare ca 0
'tiparesc indicatori despre metodele de calcul alese
Print #numarfisier, chkdif(0).Value & chkdif(1).Value & chkdif(2).Value & chkdif(3).Value '& chkdif(4).Value
Print #numarfisier, Val(sigy)
For itest = 1 To ipoints
Print #numarfisier, Val(dx(itest)) & " , " & spatiu & Val(dy(itest))
Next itest

Case "Regression"
Print #numarfisier, "regression methods"
Print #numarfisier, CStr(lst1.ListIndex + 1) & CStr(lst2.ListIndex + 1)
Print #numarfisier, CStr(txtreg.Text)
'scriu functia de conversie 3 intregi cu suma >0 si <3
Print #numarfisier, chkreg(0).Value & chkreg(1).Value & chkreg(2).Value
Print #numarfisier, CStr(Val(sigy))
For itest = 1 To ipoints
Print #numarfisier, Val(dx(itest)) & " , " & spatiu & Val(dy(itest))
Next itest




Case "CRTA"
Print #numarfisier, "CRTA methods"
Print #numarfisier, CStr(lst1.ListIndex + 1) & CStr(lst2.ListIndex + 1)
'indicator asupra metodelor alese, 1 pt on si 0 pt off, suma trebuie sa fie mai mare ca 0
'Print #numarfisier, Abs(CInt(optcrt(1).Value)) & Abs(CInt(optcrt(2).Value))
'tiparesc abaterea standard, desi nu o folosesc acum
Print #numarfisier, CStr(Val(sigy))

For i = 1 To 9
Print #numarfisier, Val(txtcrta(i - 1)) & "    ";
Next i
Print #numarfisier, ""

For itest = 1 To ipoints1
Print #numarfisier, Val(dx(itest)) & " , " & spatiu & Val(dy(itest))
Next itest
Print #numarfisier, "                              "
For itest = ipoints1 + 2 To ipoints1 + ipoints2 + 1
Print #numarfisier, Val(dx(itest)) & " , " & spatiu & Val(dy(itest))
Next itest
End Select
Close #numarfisier
Exit Sub
handleit:
Exit Sub
End Sub


Private Sub shiftx_Click()
On Error GoTo handleit
Dim puncte As Integer
Select Case tabdata.Caption
Case "CRTA"
MsgBox "Inactive option for this method."
Case Else
Call verif_grid(1, Grid1.Rows - 1, puncte)
If puncte < 7 Then MsgBox "Not enough data points.": Exit Sub
Dim lx() As Double
ReDim lx(puncte)
t# = InputBox("Insert shift value:")
Grid1.Col = 1
For i = 1 To puncte
Grid1.Row = i
Grid1.Text = CStr(Val(Grid1.Text) + t#)
Next i
End Select
Exit Sub
handleit:
MsgBox " Unexpected error."
Exit Sub
End Sub

Private Sub smooth_Click()
'face smooth in data grid cu un polinom de gradul doi, pe patru puncte;
'doua cate doua adiacente
On Error GoTo handleit
Dim puncte As Integer
Select Case tabdata.Caption
Case "CRTA"
MsgBox "Inactive option for this method."
Case Else
Call verif_grid(1, Grid1.Rows - 1, puncte)
If puncte < 7 Then MsgBox "Not enough data points.": Exit Sub
Dim lx() As Double, ly() As Double, coef() As Double
ReDim lx(puncte), ly(puncte), coef(puncte, 3)
Dim aa(5, 3) As Double, ii(5) As Double, xx(3) As Double, eroare As Boolean
For i = 1 To puncte
Grid1.Row = i
Grid1.Col = 1: lx(i) = Val(Grid1.Text)
Grid1.Col = 2: ly(i) = Val(Grid1.Text)
Next i
Grid1.Col = 2
For j = 3 To puncte - 2
aa(1, 1) = 1: aa(1, 2) = lx(j - 2): aa(1, 3) = (lx(j - 2)) ^ 2 ': aa(1, 4) = (lx(j - 2)) ^ 3
aa(2, 1) = 1: aa(2, 2) = lx(j - 1): aa(2, 3) = (lx(j - 1)) ^ 2 ': aa(2, 4) = (lx(j - 1)) ^ 3
aa(3, 1) = 1: aa(3, 2) = lx(j): aa(3, 3) = (lx(j)) ^ 2 ': aa(3, 4) = (lx(j)) ^ 3
aa(4, 1) = 1: aa(4, 2) = lx(j + 1): aa(4, 3) = (lx(j + 1)) ^ 2 ': aa(4, 4) = (lx(j + 1)) ^ 3
aa(5, 1) = 1: aa(5, 2) = lx(j + 2): aa(5, 3) = (lx(j + 2)) ^ 2 ': aa(5, 4) = (lx(j + 2)) ^ 3
ii(1) = ly(j - 2): ii(2) = ly(j - 1): ii(3) = ly(j): ii(4) = ly(j + 1): ii(5) = ly(j + 2)
Call pseudoinv(5, 3, aa(), ii(), xx(), lowsize, eroare)
If eroare Then Err.Raise 1101
coef(j, 1) = xx(1): coef(j, 2) = xx(2): coef(j, 3) = xx(3) ': coef(j, 4) = xx(4)
Next j
'Grid1.Row = 2
'Grid1.Text = Format$(Val(coef(3, 1) + coef(3, 2) * lx(2) + coef(3, 3) * lx(2) * lx(2)), "######0.0#######")
'Grid1.Row = puncte - 1
'Grid1.Text = Format$(Val(coef(puncte - 2, 1) + coef(puncte - 2, 2) * lx(puncte - 1) + coef(puncte - 2, 3) * lx(puncte - 1) * lx(puncte - 1)), "######0.0#######")

For j = 3 To puncte - 2
Grid1.Row = j
Grid1.Text = Format$(Val(coef(j, 1) + coef(j, 2) * lx(j) + coef(j, 3) * lx(j) ^ 2), "######0.0#######")
Next j


End Select
Exit Sub
handleit:
MsgBox " Unexpected error."
Exit Sub
End Sub

Private Sub swap_Click()
On Error GoTo handleit
Dim s1 As String, s2 As String
Grid1.Col = 1
For i% = 1 To Grid1.Rows - 1
Grid1.Row = i%
s1 = Grid1.Text
Grid1.Col = 2
s2 = Grid1.Text
Grid1.Text = s1
Grid1.Col = 1
Grid1.Text = s2
Next i%
Exit Sub
handleit:
Exit Sub
End Sub

Private Sub tabdata_Click(PreviousTab As Integer)
'previoustab este un intreg, incepe cu 0
'indiferent de ce scrie individul in ele, ultimul tab decide metoda
lst2_Click
Me.MousePointer = 0
End Sub

Private Sub tabdata_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
data_editor.PopupMenu mnu_file, , , , mnu_save
End If
End Sub

Private Sub txtname_Change()
nume_exp = txtname.Text
End Sub

