VERSION 4.00
Begin VB.Form main_display 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5010
   ClientLeft      =   1830
   ClientTop       =   1995
   ClientWidth     =   7680
   ForeColor       =   &H00000000&
   Height          =   5700
   HelpContextID   =   1000
   Icon            =   "versat.frx":0000
   Left            =   1770
   LinkTopic       =   "Form1"
   ScaleHeight     =   95.256
   ScaleMode       =   0  'User
   ScaleWidth      =   102.033
   Top             =   1365
   Width           =   7800
   Begin RichtextLib.RichTextBox richtxtlog 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13356
      _ExtentY        =   8488
      _StockProps     =   69
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         name            =   "Courier New"
         charset         =   0
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"versat.frx":0442
      Appearance      =   0
      RightMargin     =   50000
   End
   Begin VB.Line Line1 
      X1              =   0.598
      X2              =   159.36
      Y1              =   0
      Y2              =   0
   End
   Begin MSComDlg.CommonDialog comdialog1 
      Left            =   8040
      Top             =   2280
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CancelError     =   -1  'True
   End
   Begin VB.Menu menu_file 
      Caption         =   "&File"
      HelpContextID   =   1101
      Begin VB.Menu menu_editor 
         Caption         =   "&Notepad"
      End
      Begin VB.Menu mnuword 
         Caption         =   "&Wordpad"
      End
      Begin VB.Menu m3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_printsetup 
         Caption         =   "Printer Setup"
      End
      Begin VB.Menu menu_ 
         Caption         =   "-"
      End
      Begin VB.Menu menu_quit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnudata 
      Caption         =   "&Data"
      HelpContextID   =   1201
      Begin VB.Menu mnuopen 
         Caption         =   "&Open data file"
      End
      Begin VB.Menu mnu_3 
         Caption         =   "-"
      End
      Begin VB.Menu menu_edit_data 
         Caption         =   "&Edit data "
      End
      Begin VB.Menu mnu_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_print_data 
         Caption         =   "&Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu save_data 
         Caption         =   "&Save as"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_results 
      Caption         =   "&Results"
      HelpContextID   =   1301
      Begin VB.Menu mnufonts 
         Caption         =   "&Format"
      End
      Begin VB.Menu mnu_viewlog 
         Caption         =   "&View"
         Checked         =   -1  'True
      End
      Begin VB.Menu menu_1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_save_results 
         Caption         =   "&Save"
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_clipboard 
         Caption         =   "&Copy to Clipboard"
      End
      Begin VB.Menu menu_print_results 
         Caption         =   "Send to &WordPad"
      End
      Begin VB.Menu mnu_1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_discard 
         Caption         =   "&Discard"
      End
   End
   Begin VB.Menu menu_compute 
      Caption         =   "&Compute"
      HelpContextID   =   1401
      Begin VB.Menu menu_comp_param 
         Caption         =   "&Parameters"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_graphic 
         Caption         =   "&Graphic"
      End
      Begin VB.Menu menu_simulate 
         Caption         =   "&Simulation"
      End
   End
   Begin VB.Menu menu_help 
      Caption         =   "&Help"
      HelpContextID   =   1501
      Begin VB.Menu menu_about 
         Caption         =   "&About..."
      End
      Begin VB.Menu menu_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu_contents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu helpsearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu helponhelp 
         Caption         =   "&Help on Help"
      End
   End
End
Attribute VB_Name = "main_display"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
DoEvents
Load data_editor
initializare
Screen.MousePointer = 0
Me.Caption = nume_prog & " - main "
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
main_display.PopupMenu mnu_results
End If
End Sub
Private Sub Form_Resize()
On Error GoTo handleit
Me.Caption = nume_prog & " - main "
If Me.WindowState = 1 Then Me.Caption = nume_prog
richtxtlog.Height = main_display.ScaleHeight - 2.5
richtxtlog.Width = main_display.ScaleWidth - 1.5
line1.x2 = richtxtlog.Width + 1.5
handleit:
Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim t As Integer
t = MsgBox("   Are you sure you want to quit ?     ", 48 + 4 + 256, nume_prog)
If t = 6 Then Close: End
Cancel = 1
End Sub

Private Sub helponhelp_Click()
retval = WinHelp(main_display.hwnd, "versat10.hlp", HELP_HELPONHELP, CLng(0))
End Sub

Private Sub helpsearch_Click()
retval = WinHelp(main_display.hwnd, "versat10.hlp", HELP_KEY, CLng(0))

End Sub

Private Sub menu_about_Click()
apel = True
about.Show
End Sub
Private Sub menu_comp_param_Click()
Call compute_parameters
End Sub

Private Sub menu_contents_Click()
retval = WinHelp(main_display.hwnd, "versat10.hlp", HELP_INDEX, CLng(0))
End Sub

Private Sub menu_discard_Click()
Dim itest As Integer
itest = MsgBox(" Do you want to clear the results ?", vbYesNo + vbDefaultButton2 + vbExclamation, nume_prog)
If itest = vbYes Then
richtxtlog.Text = ""
Call scrie_log("Restarting " & nume_prog & vbCrLf & licenta & vbCrLf & Format$(Now, "dddd, mmm d yyyy") + ", " + Format$(Now, "hh:mm:ss"))
End If
End Sub
Private Sub menu_edit_data_Click()
menu_comp_param.Enabled = True
'data_editor.WindowState = 0
data_editor.Show
main_display.Hide
data_editor.SetFocus
End Sub
Private Sub menu_editor_Click()
 Shell "notepad.exe", 1
End Sub


Private Sub menu_print_data_Click()
'trec la menu de tiparire in data editor
data_editor.mnu_print_Click
End Sub
Private Sub menu_print_results_Click()
On Error GoTo handleit
richtxtlog.SaveFile "_vrsat.rtf", rtfRTF
Close
ret = Shell("write.exe _vrsat.rtf", 3)
Exit Sub
handleit:
MsgBox " Unexpected error."
Exit Sub
End Sub

Private Sub menu_quit_Click()
t% = MsgBox("  Are you sure you want to quit ?     ", 48 + 4 + 256, nume_prog)
If t% = 6 Then Close: retval = WinHelp(hwnd, dummy$, HELP_QUIT, 0): End
End Sub

Private Sub menu_save_results_Click()
On Error GoTo handleit:
Dim numarfisier As Integer, itest As Integer
numarfisier = FreeFile
Call deschide_fisier(CurDir$ + "\", 2, " text file (*.txt) |*.txt| results file (*.out) |*.out| show all (*.*) |*.*", 2)
richtxtlog.SaveFile outputfile, rtfText
Exit Sub
handleit:
If Not (Err.Number = 32755) Then MsgBox "Unexpected error." & vbCrLf & CStr(Err.Description)
Exit Sub
End Sub

Private Sub menu_simulate_Click()
gindicator(2) = False
parameters.Show
main_display.Hide
End Sub

Private Sub mnu_clipboard_Click()
On Error GoTo handleit
Clipboard.SetText richtxtlog.Text
Exit Sub
handleit:
MsgBox " Unexpected error. "
Exit Sub
End Sub

Private Sub mnu_graphic_Click()
'aceasta este rutina pentru grafic
'se pot vedea punctele experimentale, patru metode
Dim eroare As Boolean
'Call verif_data(eroare)
'If eroare Then Exit Sub
graphics.mnu_view_exp.Enabled = gindicator(1)
graphics.txtr(0).Text = CStr(xstart)
graphics.txtr(1).Text = CStr(xend)
graphics.txtr(2).Text = CStr(ystart)
graphics.txtr(3).Text = CStr(yend)
graphics.Visible = True
End Sub



Private Sub mnu_printsetup_Click()
On Error GoTo handleit
main_display.comdialog1.Flags = cdlPDPrintSetup
'main_display.comdialog1.Flags =  cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoWarning Or cdlPDUseDevModeCopies Or cdlPDAllPages Or cdlPDNoSelection
main_display.comdialog1.ShowPrinter
Exit Sub
handleit:
Exit Sub
End Sub



Private Sub mnu_viewlog_Click()
mnu_viewlog.Checked = Not (mnu_viewlog.Checked)
richtxtlog.Visible = mnu_viewlog.Checked
End Sub

Private Sub mnufonts_Click()
On Error GoTo handleit
comdialog1.Flags = cdlCFBoth Or &H100&
comdialog1.ShowFont
With richtxtlog
.SelFontName = comdialog1.FontName
.SelFontSize = comdialog1.FontSize
.SelBold = comdialog1.FontBold
.SelItalic = comdialog1.FontItalic
.SelStrikethru = comdialog1.FontStrikethru
.SelUnderline = comdialog1.FontUnderline
.SelColor = comdialog1.Color
End With
Exit Sub
handleit:
If Err.Number = 32755 Then Exit Sub
MsgBox CStr(Err.Number) & vbCrLf & " Unexpected error." & Err.Description
Exit Sub
End Sub

Private Sub mnuopen_Click()
menu_comp_param.Enabled = True
data_editor.Show
data_editor.Refresh
main_display.Hide
data_editor.mnu_open_Click
End Sub

Private Sub mnuword_Click()
On Error GoTo handle
ret = Shell("write.exe", 3)
Exit Sub
handle:
MsgBox "Error trying to open the wordpad program."
Exit Sub
End Sub

Private Sub save_data_Click()
data_editor.mnu_save_Click
End Sub

Private Sub richtxtlog_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
main_display.PopupMenu mnu_results
End If

End Sub

