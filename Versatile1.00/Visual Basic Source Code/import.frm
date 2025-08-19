VERSION 4.00
Begin VB.Form import_form 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Import data"
   ClientHeight    =   5910
   ClientLeft      =   1710
   ClientTop       =   810
   ClientWidth     =   8865
   Height          =   6600
   HelpContextID   =   5000
   Icon            =   "import.frx":0000
   Left            =   1650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8865
   Top             =   180
   Width           =   8985
   Begin VB.TextBox txtscale 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   5
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox txtscale 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox txtscale 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtscale 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add values"
      Enabled         =   0   'False
      Height          =   372
      HelpContextID   =   5010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton cmdsetscale 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Set scale"
      Height          =   372
      HelpContextID   =   5020
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1332
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label actual 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   238
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label actual 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   238
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Y (2)"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   13
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "X (2)"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Y (1)"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "X (1)"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image pctimp 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5172
      Left            =   1560
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7212
   End
   Begin VB.Label ly 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y="
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   238
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lx 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X="
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   238
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label lbly 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "undefined"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblx 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "undefined"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu mnu_open 
         Caption         =   "&Open Image"
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "import_form"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim formstdheight, formstdwidth, minformwidth, minformheight
Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single
Dim wait As Boolean, minx As Single, miny As Single, maxx As Single, maxy As Single, setscala As Boolean


Private Sub cmdadd_Click()
On Error GoTo handleit
If Val(Right$(lx.Caption, (Len(lx.Caption) - 2))) < 0 Or Val(Right$(ly.Caption, (Len(ly.Caption) - 2))) < 0 Then Exit Sub
data_editor.Grid1.Rows = data_editor.Grid1.Rows + 1
data_editor.Grid1.Col = 1
data_editor.Grid1.Text = Right$(lx.Caption, (Len(lx.Caption) - 2))
data_editor.Grid1.Col = 2
data_editor.Grid1.Text = Right$(ly.Caption, (Len(ly.Caption) - 2))
data_editor.Grid1.Row = data_editor.Grid1.Row + 1
Exit Sub
handleit:
Exit Sub
End Sub



Private Sub cmdsetscale_Click(Index As Integer)
Static i As Boolean
If i = False Then
t% = MsgBox("In order to compute a user scale: click twice on the opposite corners of a rectangle (recommended top-left and low-right) and then insert the coorresponding values for your scale.)", vbOKCancel + vbInformation, nume_prog)
If t% = vbCancel Then Exit Sub
i = True
End If
'setx = True: sety = False
setscala = True
lblx.Caption = ""
lbly.Caption = ""
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
setscala = False
Me.Caption = nume_prog & " - Import data"
End Sub



Private Sub Form_Resize()
On Error GoTo handleit
Me.Caption = nume_prog & " - Import data"
If Me.WindowState = 1 Then Me.Caption = " Import..."
For i% = 0 To 3: txtscale(i%).Text = "": Next i%
pctimp.Width = Me.Width - 3.5 * lx.Left - lx.Width
pctimp.Height = Me.Height - 8.5 * pctimp.Top
lx.Caption = "": ly.Caption = ""
actual(0).Caption = "": actual(1).Caption = ""
lblx.Caption = "undefined": lbly.Caption = "undefined"

Exit Sub
handleit:
Exit Sub
End Sub

Private Sub mnu_exit_Click()
Unload Me
End Sub

Private Sub mnu_help_Click()
retval = WinHelp(import_form.hwnd, "versat10.hlp", HELP_KEY, CLng(0))
End Sub

 Sub mnu_open_Click()
On Error GoTo handleit
Dim numarfisier As Integer, stest As String, dtest(4) As Double, itest As Integer, i As Integer
numarfisier = FreeFile
Call deschide_fisier(CurDir$ + "\", 1, " bmp file (*.bmp) |*.bmp| wmf file (*.wmf) |*.wmf| show all (*.*) |*.*", 1)
If no_input Then Exit Sub
pctimp.Picture = LoadPicture(inputfile)
'importform!cmdsetx.Visible = True: cmdsety.Enabled = True
'txtxscale.Enabled = True: txtyscale.Enabled = True
Exit Sub
handleit:
If Err.Number = 62 Then Exit Sub
MsgBox "Error number " & CStr(Err.Number) & vbCrLf & CStr(Err.Description) & vbCrLf & "Unexpected error. Check the data or report the conditions to the author."
'cmdsetx.Enabled = False: cmdsety.Enabled = False
'txtxscale.Enabled = False: txtyscale.Enabled = False
Exit Sub

End Sub

Private Sub pctimp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo handle
If Abs(Val(txtscale(2).Text) - Val(txtscale(0).Text)) > 0 And Abs(Val(txtscale(3).Text) - Val(txtscale(1).Text)) > 0 Then
    actual(0).Caption = "X=" & Format$(Val(txtscale(0).Text) + ((x - x1) / (x2 - x1) * (Val(txtscale(2).Text) - Val(txtscale(0).Text))), "####0.000")
    actual(1).Caption = "Y=" & Format$(Val(txtscale(1).Text) + ((y - y1) / (y2 - y1) * (Val(txtscale(3).Text) - Val(txtscale(1).Text))), "####0.000")
cmdadd.Enabled = True
Else
cmdadd.Enabled = False
End If
Label.Caption = CStr(x) & ", " & CStr(y)
Exit Sub
handle:
Exit Sub
End Sub

Private Sub pctimp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo handle
      If setscala Then
        If (lblx.Caption = "") Then
            lblx.Caption = Format$(x, "####0") & "-": x1 = x
            lblx.Caption = lblx.Caption & Format$(y, "####0"): y1 = y
        Exit Sub
        End If
       lbly.Caption = Format$(x, "####0") & "-": x2 = x
       lbly.Caption = lbly.Caption & Format$(y, "####0"): y2 = y
       setscala = False
       End If
'inseamna ca e right button, modific x si y
'If Val(txtxscale) < 0.001 Or Val(txtyscale) < 0.001 Then Exit Sub
If Abs(Val(txtscale(2).Text) - Val(txtscale(0).Text)) > 0 And Abs(Val(txtscale(3).Text) - Val(txtscale(1).Text)) > 0 Then
    lx.Caption = "X=" & Format$(Val(txtscale(0).Text) + ((x - x1) / (x2 - x1) * (Val(txtscale(2).Text) - Val(txtscale(0).Text))), "####0.000")
    ly.Caption = "Y=" & Format$(Val(txtscale(1).Text) + ((y - y1) / (y2 - y1) * (Val(txtscale(3).Text) - Val(txtscale(1).Text))), "####0.000")
'ly.Caption = "Y=" & Format$(Y / (maxy - miny) * Val(txtyscale), "#####0.000")
End If
Exit Sub
handle:
Exit Sub
End Sub







Private Sub txtscale_Change(Index As Integer)
If Val(Val(txtscale(2).Text) - Val(txtscale(0).Text)) > 0.1 And Val(Val(txtscale(3).Text) - Val(txtscale(1).Text)) > 0.1 Then cmdadd.Enabled = True
End Sub

Private Sub txtscale_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
