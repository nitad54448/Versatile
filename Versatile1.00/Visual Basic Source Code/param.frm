VERSION 4.00
Begin VB.Form parameters 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Simulate"
   ClientHeight    =   3792
   ClientLeft      =   1116
   ClientTop       =   996
   ClientWidth     =   5748
   ForeColor       =   &H00000000&
   Height          =   4344
   HelpContextID   =   4000
   Icon            =   "param.frx":0000
   Left            =   1068
   LinkTopic       =   "Form1"
   ScaleHeight     =   3792
   ScaleWidth      =   5748
   Top             =   492
   Width           =   5844
   Begin VB.CheckBox chk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&CRTA"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   1332
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Height          =   288
      HelpContextID   =   4030
      Index           =   7
      Left            =   2520
      TabIndex        =   27
      Text            =   "0.0"
      Top             =   480
      Width           =   1092
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Height          =   288
      HelpContextID   =   4090
      Index           =   6
      Left            =   2520
      TabIndex        =   6
      Text            =   "5"
      Top             =   3120
      Width           =   1092
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Height          =   288
      HelpContextID   =   4080
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Text            =   "800"
      Top             =   2640
      Width           =   1092
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Height          =   288
      HelpContextID   =   4070
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Text            =   "300"
      Top             =   2280
      Width           =   1092
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Height          =   288
      HelpContextID   =   4060
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Text            =   "5"
      Top             =   1800
      Width           =   1092
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Height          =   288
      HelpContextID   =   4050
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Text            =   "120.0"
      Top             =   1320
      Width           =   1092
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Height          =   288
      HelpContextID   =   4040
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Text            =   "10E12"
      Top             =   960
      Width           =   1092
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
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
      Height          =   288
      HelpContextID   =   4020
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Text            =   "1.0"
      Top             =   120
      Width           =   1092
   End
   Begin Spin.SpinButton spin 
      Height          =   276
      Index           =   7
      Left            =   3600
      TabIndex        =   28
      Top             =   480
      Width           =   228
      _Version        =   65536
      _ExtentX        =   402
      _ExtentY        =   487
      _StockProps     =   73
      ForeColor       =   -2147483630
      Enabled         =   0   'False
      BorderColor     =   8421504
      BorderThickness =   0
      ShadeColor      =   16
      ShadowBackColor =   5
      ShadowForeColor =   18
      TdThickness     =   1
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      Enabled         =   0   'False
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
      Height          =   252
      Index           =   7
      Left            =   1800
      TabIndex        =   26
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   612
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select change"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   7.8
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   25
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "0.30%"
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
      Height          =   192
      Index           =   2
      Left            =   4848
      TabIndex        =   24
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "20"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   23
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "0.05"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   22
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin ComctlLib.Slider slider 
      Height          =   3135
      HelpContextID   =   4100
      Left            =   4200
      TabIndex        =   7
      Top             =   360
      Width           =   480
      _Version        =   65536
      Orientation     =   1
      _ExtentX        =   847
      _ExtentY        =   5525
      _StockProps     =   64
      LargeChange     =   20
      Max             =   400
      Min             =   1
      SelStart        =   6
      TickFrequency   =   20
      Value           =   6
   End
   Begin Spin.SpinButton spin 
      Height          =   276
      Index           =   6
      Left            =   3600
      TabIndex        =   14
      Top             =   3120
      Width           =   228
      _Version        =   65536
      _ExtentX        =   402
      _ExtentY        =   487
      _StockProps     =   73
      ForeColor       =   -2147483630
      BorderColor     =   8421504
      BorderThickness =   0
      ShadeColor      =   16
      ShadowBackColor =   5
      ShadowForeColor =   18
      TdThickness     =   1
   End
   Begin Spin.SpinButton spin 
      Height          =   276
      Index           =   5
      Left            =   3600
      TabIndex        =   13
      Top             =   2640
      Width           =   228
      _Version        =   65536
      _ExtentX        =   402
      _ExtentY        =   487
      _StockProps     =   73
      ForeColor       =   -2147483630
      BorderColor     =   8421504
      BorderThickness =   0
      ShadeColor      =   16
      ShadowBackColor =   5
      ShadowForeColor =   18
      TdThickness     =   1
   End
   Begin Spin.SpinButton spin 
      Height          =   276
      Index           =   4
      Left            =   3600
      TabIndex        =   12
      Top             =   2280
      Width           =   228
      _Version        =   65536
      _ExtentX        =   402
      _ExtentY        =   487
      _StockProps     =   73
      ForeColor       =   -2147483630
      BorderColor     =   8421504
      BorderThickness =   0
      ShadeColor      =   16
      ShadowBackColor =   5
      ShadowForeColor =   18
      TdThickness     =   1
   End
   Begin Spin.SpinButton spin 
      Height          =   276
      Index           =   3
      Left            =   3600
      TabIndex        =   11
      Top             =   1800
      Width           =   228
      _Version        =   65536
      _ExtentX        =   402
      _ExtentY        =   487
      _StockProps     =   73
      ForeColor       =   -2147483630
      BorderColor     =   8421504
      BorderThickness =   0
      ShadeColor      =   16
      ShadowBackColor =   5
      ShadowForeColor =   18
      TdThickness     =   1
   End
   Begin Spin.SpinButton spin 
      Height          =   276
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Top             =   1320
      Width           =   228
      _Version        =   65536
      _ExtentX        =   402
      _ExtentY        =   487
      _StockProps     =   73
      ForeColor       =   -2147483630
      BorderColor     =   8421504
      BorderThickness =   0
      ShadeColor      =   16
      ShadowBackColor =   5
      ShadowForeColor =   18
      TdThickness     =   1
   End
   Begin Spin.SpinButton spin 
      Height          =   276
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   960
      Width           =   228
      _Version        =   65536
      _ExtentX        =   402
      _ExtentY        =   487
      _StockProps     =   73
      ForeColor       =   -2147483630
      BorderColor     =   8421504
      BorderThickness =   0
      ShadeColor      =   16
      ShadowBackColor =   5
      ShadowForeColor =   18
      TdThickness     =   1
   End
   Begin Spin.SpinButton spin 
      Height          =   276
      Index           =   0
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   228
      _Version        =   65536
      _ExtentX        =   402
      _ExtentY        =   487
      _StockProps     =   73
      ForeColor       =   -2147483630
      BorderColor     =   8421504
      BorderThickness =   0
      ShadeColor      =   16
      ShadowBackColor =   5
      ShadowForeColor =   18
      TdThickness     =   1
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of p(x) terms"
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
      Height          =   252
      Index           =   6
      Left            =   0
      TabIndex        =   21
      Top             =   3120
      UseMnemonic     =   0   'False
      Width           =   2412
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final temperature /K"
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
      Height          =   252
      Index           =   5
      Left            =   0
      TabIndex        =   20
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   2412
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Initial temperature /K"
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
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   2292
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Heating rate (K/min)"
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
      Height          =   252
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   2412
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Activation energy (kJ/mol)"
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
      Height          =   252
      Index           =   2
      Left            =   0
      TabIndex        =   17
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   2412
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exponential factor (1/sec)"
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
      Height          =   252
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   2412
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "n"
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
      Height          =   252
      Index           =   0
      Left            =   1800
      TabIndex        =   15
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   612
   End
   Begin VB.Menu ok 
      Caption         =   "&OK"
   End
   Begin VB.Menu close 
      Caption         =   "&Close"
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "parameters"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Sub simulatecrta(a1() As Double, t1() As Double, n As Double, m As Double, a As Double, E As Double, c As Double, eroare As Boolean)
Dim i As Integer, eps As Single
On Error GoTo handle
eps = 0.00001
If n < eps Then n = 0
If m < eps Then m = 0
Select Case m
Case 0
    Select Case n
    Case 0
    For i = 2 To 201
    a1(i - 1) = (i - 1) / 201
    t1(i - 1) = E / 8.314 * 1 / (Log(a / c))
    Next i
    Case Else
    For i = 2 To 201
    a1(i - 1) = (i - 1) / 201
    t1(i - 1) = E / 8.314 * 1 / (Log(a * ((1 - a1(i - 1)) ^ n) / c))
    Next i
    End Select
Case Else
    Select Case n
    Case 0
    For i = 2 To 201
    a1(i - 1) = (i - 1) / 201
    t1(i - 1) = E / 8.314 * 1 / (Log(a * (a1(i - 1) ^ m) / c))
    Next i
    Case Else
    For i = 2 To 201
    a1(i - 1) = (i - 1) / 201
    t1(i - 1) = E / 8.314 / (Log(a / c * (a1(i - 1) ^ m) * ((1 - a1(i - 1)) ^ n)))
    Next i
    End Select
End Select
eroare = False
Exit Sub
handle:
eroare = True
Exit Sub
End Sub


Private Sub chk_Click()
If (chk.Value) Then
lbl1(7).Enabled = True
txt(7).Enabled = True
Spin(7).Enabled = True
lbl1(3).Caption = "Decomposition rate (1/sec)"
lbl1(4).Enabled = False: txt(4).Enabled = False: Spin(4).Enabled = False
lbl1(5).Enabled = False: txt(5).Enabled = False: Spin(5).Enabled = False
lbl1(6).Enabled = False: txt(6).Enabled = False: Spin(6).Enabled = False
Else
lbl1(4).Enabled = True: txt(4).Enabled = True: Spin(4).Enabled = True
lbl1(5).Enabled = True: txt(5).Enabled = True: Spin(5).Enabled = True
lbl1(6).Enabled = True: txt(6).Enabled = True: Spin(6).Enabled = True
lbl1(7).Enabled = False
txt(7).Enabled = False
Spin(7).Enabled = False
lbl1(3).Caption = "Heating rate (K/min)"

End If

End Sub

Sub close_Click()
If Not (apelsimulare) Then main_display.Show
apelsimulare = False
Me.Hide

End Sub







Private Sub Form_Load()
Me.Caption = nume_prog & " - Simulation"
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Slider.Value = 6
End Sub

Private Sub Form_Resize()
Me.Caption = nume_prog & " - Simulation"
If Me.WindowState = 1 Then Me.Caption = " Simulation."

End Sub

Private Sub Form_Unload(Cancel As Integer)
close_Click
End Sub

Private Sub help_Click()
retval = WinHelp(parameters.hwnd, "versat10.hlp", HELP_KEY, CLng(0))

End Sub

Private Sub ok_Click()
Me.MousePointer = 11
On Error GoTo handleit
For i% = 0 To 7
If Len(txt(i%).Text) < 1 Then Err.Raise 1101, , "A text field is empty; number " & CStr(i% + 1) & ". Correct this and try again."
Next i%
If CDbl(txt(0).Text) < 0 Or CDbl(txt(0).Text) > 3 Then Err.Raise 1101, , "Incorrect value of reaction order."
If Val(txt(1).Text) < 0 Or Val(txt(1).Text) > 1E+80 Then Err.Raise 1101, , "Incorrect value of preexponential factor."
If 1000# * CDbl(txt(2).Text) < 10# Or 1000# * CDbl(txt(2).Text) > 500000# Then Err.Raise 1101, , "Incorrect value of activation energy."
If CDbl(txt(3).Text) < 0.0000001 Or CDbl(txt(3).Text) > 100 Then Err.Raise 1101, , "Incorrect value of heating or decomposition rate."
If CDbl(txt(4).Text) < 100 Or CDbl(txt(4).Text) > 1500 Then Err.Raise 1101, , "Incorrect value of initial temperature."
If CDbl(txt(5).Text) < CDbl(txt(4).Text) Or CDbl(txt(5).Text) > 1600 Then Err.Raise 1101, , "Incorrect value of final temperature."
If CInt(txt(6).Text) < 1 Or CInt(txt(6).Text) > 15 Then Err.Raise 1101, , "Incorrect value of p(x) terms. Must be a positive integer and lower than E/RT."
'(al() As Double, te() As Double, tempstart As Double, tempend As Double, n As Double, a As Double, e As Double, r As Double, npx As Integer, eroare As Boolean)
Dim al(200) As Double, te(200) As Double, eroare As Boolean
If (chk.Value = 0) Then
Call simulate(al(), te(), CDbl(txt(4).Text), CDbl(txt(5).Text), CDbl(txt(0).Text), CDbl(txt(1).Text), 1000# * CDbl(txt(2).Text), CDbl(txt(3).Text) / 60#, CInt(txt(6).Text), eroare)
If eroare Then Err.Raise 1101, , "Unexpected error. Check your data."
For j% = 1 To 200
xgraf(j%, 2) = te(j%): ygraf(j%, 2) = al(j%)
Next j%

Else
''simulatecrta
Call simulatecrta(al(), te(), Val(txt(0).Text), Val(txt(7).Text), Val(txt(1).Text), 1000 * Val(txt(2).Text), Val(txt(3).Text), eroare)
If eroare Then Err.Raise 1101, , "Unexpected error. Check your data."
For j% = 1 To 200
xgraf(j%, 2) = al(j%): ygraf(j%, 2) = te(j%)
Next j%
End If

If Not (apelsimulare) Then
scrie_log linie
If chk.Value = 0 Then
scrie_log "Simulated (T,alpha), for:" & vbCrLf & "Reaction order = " & CStr(CDbl(txt(0).Text)) & vbCrLf & "Preexponential = " & Format(CDbl(txt(1).Text), "0.000E+00") & vbCrLf & "Activation energy = " & CStr(1000# * CDbl(txt(2).Text)) & vbCrLf & "Heating rate (K/min) = " & CStr(CDbl(txt(3).Text)) & vbCrLf & "P(x) terms = " & CStr(CInt(txt(6).Text)) & vbCrLf & linie
scrie_log "Temp. /K" & " " & Chr(9) & "conversion"

Else
scrie_log "Simulated CRTA data, for:" & vbCrLf & " n= " & CStr(CDbl(txt(0).Text)) & vbCrLf & " m= " & CStr(CDbl(txt(7).Text)) & vbCrLf & "Preexponential = " & Format(CDbl(txt(1).Text), "0.000E+00") & vbCrLf & "Activation energy = " & CStr(1000# * CDbl(txt(2).Text)) & vbCrLf & "Decomposition rate (1/sec) = " & CStr(CDbl(txt(3).Text)) & vbCrLf & linie
scrie_log "conversion" & " " & Chr(9) & " Temp /K"

End If
For j% = 1 To 200
scrie_log Format$(xgraf(j%, 2), "###0.000") & "   " & Chr(9) & Format$(ygraf(j%, 2), "0.000000")
Next j%
scrie_log linie
End If
gindicator(2) = True
Err.Clear
If apelsimulare Then Me.MousePointer = 0: graphics.redeseneaza: Exit Sub
main_display.Show
Me.MousePointer = 0
apelsimulare = False
Unload Me
Exit Sub
handleit:
MsgBox Err.Description, vbOKOnly + vbInformation, nume_prog
gindicator(2) = False
'If apelsimulare Then Me.MousePointer = 0: Exit Sub
'main_display.Show
Me.MousePointer = 0
'Unload Me
'apelsimulare = False
Exit Sub

End Sub

Private Sub slider_Change()
Label1(2).Caption = Format$(((Slider.Value / 20)), "#0.00") & " % "
End Sub

Private Sub Slider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
slider_Change
End Sub

Private Sub Spin_SpinDown(Index As Integer)
On Error GoTo handle
Select Case Index
Case 0
txt(Index).Text = Format$((Val(txt(Index).Text) - Slider.Value * Val(txt(Index).Text) / 2000), "0.0000")
If Val(txt(Index).Text) < 0 Then txt(Index).Text = "0.0000"
Case 1
txt(Index).Text = Format$((Val(txt(Index).Text) - Slider.Value * Val(txt(Index).Text) / 2000), "0.0000E+00")
If Val(txt(Index).Text) < 100 Then txt(Index).Text = "100.00"
Case 2
txt(Index).Text = Format$((Val(txt(Index).Text) - Slider.Value * Val(txt(Index).Text) / 2000), "#00.0000")
If Val(txt(Index).Text) < 10 Then txt(Index).Text = "10.0000"
Case 3
txt(Index).Text = Format$((Val(txt(Index).Text) - Slider.Value * Val(txt(Index).Text) / 2000), "0.0000E+00")
If Val(txt(Index).Text) < 0.0000001 Then txt(Index).Text = "0.0000001"
Case 4
txt(Index).Text = Format$((Val(txt(Index).Text) - Slider.Value * Val(txt(Index).Text) / 2000), "#000.00")
If Val(txt(Index).Text) < 100 Then txt(Index).Text = "100.00"
Case 5
txt(Index).Text = Format$((Val(txt(Index).Text) - Slider.Value * Val(txt(Index).Text) / 2000), "#000.00")
If Val(txt(Index).Text) < Val(txt(4).Text) Then txt(Index).Text = Val(txt(4).Text) + "100.00"
Case 6
txt(Index).Text = Format$(Val(txt(Index).Text - 1), "#0")
If Val(txt(Index).Text) < 1 Then txt(Index).Text = "1"
Case 7
txt(Index).Text = Format$((Val(txt(Index).Text) - Slider.Value * Val(txt(Index).Text) / 2000), "#000.00")
If Val(txt(Index).Text) < 0# Then txt(Index).Text = "0.0000"
End Select
If apelsimulare Then ok_Click
Exit Sub
handle:
Exit Sub
End Sub

Private Sub Spin_SpinUp(Index As Integer)
On Error GoTo handleit
Select Case Index
Case 0
txt(Index).Text = Format$((Val(txt(Index).Text) + Slider.Value * Val(txt(Index).Text) / 2000), "0.0000")
If Val(txt(Index).Text) > 3 Then txt(Index).Text = "3.0000"
Case 1
txt(Index).Text = Format$((Val(txt(Index).Text) + Slider.Value * Val(txt(Index).Text) / 2000), "0.0000E+00")
If Val(txt(Index).Text) > 10 ^ 80 Then txt(Index).Text = "1.0000E+80"
Case 2
txt(Index).Text = Format$((Val(txt(Index).Text) + Slider.Value * Val(txt(Index).Text) / 2000), "#00.0000")
If Val(txt(Index).Text) > 600 Then txt(Index).Text = "600.0000"
Case 3
txt(Index).Text = Format$((Val(txt(Index).Text) + Slider.Value * Val(txt(Index).Text) / 2000), "0.0000E+00")
If Val(txt(Index).Text) > 100 Then txt(Index).Text = "100.00"
Case 4
txt(Index).Text = Format$((Val(txt(Index).Text) + Slider.Value * Val(txt(Index).Text) / 2000), "#000.00")
If Val(txt(Index).Text) > 1300 Then txt(Index).Text = "1300.00"
Case 5
txt(Index).Text = Format$((Val(txt(Index).Text) + Slider.Value * Val(txt(Index).Text) / 2000), "#000.00")
If Val(txt(Index).Text) > 1600 Then txt(Index).Text = "1600.00"
Case 6
txt(Index).Text = Format$(Val(txt(Index).Text + 1), "#0")
If Val(txt(Index).Text) > 15 Then txt(Index).Text = "15"
Case 7
txt(Index).Text = Format$((Val(txt(Index).Text) - Slider.Value * Val(txt(Index).Text) / 2000), "#000.00")
If Val(txt(Index).Text) > 3# Then txt(Index).Text = "3.0000"
End Select
If apelsimulare Then ok_Click
Exit Sub
handleit:
Exit Sub
End Sub

