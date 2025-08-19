Attribute VB_Name = "Module1"
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HELP_COMMAND = &H102&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1          '  Display topic in ulTopic
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_HELPONHELP = &H4       '  Display help on using help
Public Const HELP_INDEX = &H3            '  Display index
Public Const HELP_KEY = &H101            '  Display topic for keyword in offabData
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_QUIT = &H2             '  Terminate help
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_SETINDEX = &H5         '  Set current Index for multi index help
Public Const HELP_SETWINPOS = &H203&

Option Base 1 'sirul incepe de la 1
Global apelsimulare As Boolean
Global inputfile As String, outputfile As String
Global Const nume_prog As String = " Versatile 1.00 "
'Global Const vbcrlf  As String = vbCrLf & "  "
'Global Const spatiu As String = "    "
Global licenta As String
Global nume_exp As String
Global apel As Boolean 'pentru apelul lui about
Public Const linie As String = "--------------------------------------------------------------"
Public Const spatiu As String = "    "
Global no_input As Boolean, no_output As Boolean
Global ninitial As Single, nfinal As Single, nstep As Single, rate As Single
Global ipoints As Integer
Global ipoints1 As Integer, ipoints2 As Integer 'numarul de puncte din fiecare set pentru icar
Global rate1 As Double, rate2 As Double 'vitezele de la icar2
'Global cc1 As Double, cc2 As Double
Global ivm1 As Double, ivm2 As Double, iivm1 As Double, iivm2 As Double
Global grpol As Integer, excl As Double 'gradul polinomului de interpolare+1
Global mmval() As Double, nnval() As Double, zzz() As Double 'n si m din icar
Global s() As Double, coef() As Double, nrint As Integer
Global dtempk() As Double, dalpha() As Double
Global dx() As Double, dy() As Double, ddalpha() As Double
Global Const lowsize As Double = 1E-24
Global a1min As Double, a1max As Double, na1min As Integer, na1max As Integer 'pentru icar1, atentie a1min> a1max
Global sigy As Double 'eroarea standard pentru un punct
Global xgraf(200, 7) As Double, ygraf(200, 7) As Double
Global gpoints(7) As Integer, xstart As Double, ystart As Double, xend As Double, yend As Double
Global gindicator(7) As Boolean


Function factorial(x As Integer) As Double
On Error GoTo handle
factorial = 1
For i% = 1 To x
factorial = factorial * i%
Next i%
Exit Function
handle:
Err.Clear
Exit Function
End Function

Sub initializare()
For i% = 1 To 7
gindicator(i%) = False
Next i%
End Sub

Function igrec(c0, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, x) As Double
'grad maxim zece
On Error GoTo handle
igrec = c0 + c1 * x + c2 * x ^ 2 + c3 * x ^ 3 + c4 * x ^ 4 + c5 * x ^ 5 + c6 * x ^ 6 + c7 * x ^ 7 + c8 * x ^ 8 + c9 * x ^ 9 + c10 * x ^ 10
Exit Function
handle:
Err.Clear
Exit Function
End Function

Sub compute_parameters()
On Error GoTo localhandle:
main_display.richtxtlog.Visible = False
Dim eroare As Boolean, txt As String
main_display.MousePointer = 11
scrie_log linie & vbCrLf & linie & vbCrLf & "starting..." & vbCrLf & Format$(Now, "dddd, mmm d yyyy") + ", " + Format$(Now, "hh:mm:ss") & vbCrLf & "checking data..."
For j = 3 To 6
gindicator(j) = False
Next j
Call verif_data(eroare)
'in rutina de verificat datele,ordonez punctele, numar cate perechi am,..
If eroare Then Err.Raise 1101, , "Error in data."
'aduc toate elementele la 0
Dim i As Integer, itest As Integer, panta As Double, ordonata As Double, corrs As Double, deltab As Double, deltam As Double
scrie_log linie & vbCrLf & "Experiment name : " & nume_exp & vbCrLf & linie & vbCrLf & "Loaded data : " & data_editor.lst1.Text & ",  " & data_editor.lst2.Text & vbCrLf & linie

Select Case data_editor!tabdata.Caption

Case "Integral"
scrie_log "Integral methods. "
scrie_log "Selected procedures:"
For itest% = 0 To 3
If data_editor.chkint(itest%).Value = 1 Then scrie_log "         -" & data_editor.chkint(itest%).Caption
Next itest%
scrie_log "Initial reaction order: " & data_editor.txtint(0).Text
scrie_log "Final reaction order: " & data_editor.txtint(1).Text
scrie_log "Reaction order step: " & data_editor.txtint(2).Text
scrie_log "Heating rate, K/min: " & data_editor.txtint(3).Text
scrie_log "Estimated error in conversion: " & Format$(sigy, "#0.00#") & " %"
scrie_log linie
ReDim dalpha(ipoints), ddalpha(ipoints)
        Select Case data_editor.lst2.ListIndex
        Case 0 ' alpha
            For i = 1 To ipoints: dalpha(i) = dy(i): Next i
        Case 1 ' dta,  treci la integrare
            Call surfint(ipoints, dtempk(), dy(), dalpha(), eroare)
            If eroare Then Err.Raise 1101, , "An errror encountered when trying to integrate the DTA curve."
        Case 2 ' dtg, integrare
             For j = 2 To ipoints
                stotal# = 0
                stotal# = stotal# + ((dy(j) + dy(j - 1)) / 2) * (dtempk(j) - dtempk(j - 1))
                Next j
            Call surfint(ipoints, dtempk(), dy(), dalpha(), eroare)
            If eroare Then Err.Raise 1101, , "An errror encountered when trying to integrate the DTG curve."
        Case 3 ' tg, calculez alpha
            dalpha(1) = 0
            dalpha(ipoints) = 1
            For i = 2 To ipoints - 1
                dalpha(i) = (dy(i) - dy(1)) / (dy(ipoints) - dy(1))
            Next i
        Case Else
            Err.Raise 1101, , "Unexpected error. Please report error in lst2.listindex comp_param procedure."
        End Select
    txt = "    temp./C    temp./K  " & data_editor!lst2.Text & "  [alpha]  " & vbCrLf
    For i = 1 To ipoints
    txt = txt & vbCrLf & CStr(i) & spatiu & Format$((dtempk(i) - 273.15), "#00.00") & spatiu & Format$(dtempk(i), "#000.00") & spatiu & Format$(dy(i), "####0.00###") & spatiu & Format$(dalpha(i), "0.0000")
    Next i
    scrie_log txt
    If (data_editor.chkint(0).Value = 1) Then 'coats redfern
    Call coats_redfern
    End If
   
    If (data_editor!chkint(1).Value = 1) Then   'aici este flynn
    Call flynn_wall
    End If
    
    If (data_editor!chkint(2).Value = 1) Then   'face calculul van krevelen
    Call van_krevelen
    End If
    
    If (data_editor!chkint(3).Value = 1) Then 'urbanovici
    Call urbanovici
    End If

Case "Differential"
    ReDim linx(ipoints) As Double, liny(ipoints) As Double 'variabila locala
    ReDim dalpha(ipoints), ddalpha(ipoints)
        
scrie_log "Differential methods "
scrie_log "Selected procedures:"
For itest% = 0 To 3
If data_editor.chkdif(itest%).Value = 1 Then scrie_log "         -" & data_editor.chkdif(itest%).Caption
Next itest%
scrie_log "Initial reaction order: " & data_editor.txtdif(0).Text
scrie_log "Final reaction order: " & data_editor.txtdif(1).Text
scrie_log "Reaction order step: " & data_editor.txtdif(2).Text
scrie_log "Heating rate, K/min: " & data_editor.txtdif(3).Text
scrie_log "Estimated error in conversion: " & Format$(sigy, "#0.00#") & " %"
scrie_log linie
        Select Case data_editor.lst2.ListIndex
        Case 0 ' alpha
            For i = 1 To ipoints: dalpha(i) = dy(i): Next i
            Call deriv(ipoints, dalpha(), ddalpha(), eroare)
            If eroare Then Err.Raise 1101, , "Error trying to derivate the conversion degree. Check your data."
        Case 1 ' dta,  treci la integrare' e a lui mitica asta
            Call surfint(ipoints, dtempk(), dy(), dalpha(), eroare)
            If eroare Then Err.Raise 1101, , "An errror encountered when trying to integrate the DTA curve."
            Call deriv(ipoints, dalpha(), ddalpha(), eroare)
            If eroare Then Err.Raise 1101, , "Error trying to derivate the DTA curve. Check your data."
        Case 2 ' dtg, integrare
             For j = 2 To ipoints
                stotal# = 0
                stotal# = stotal# + ((dy(j) + dy(j - 1)) / 2) * (dtempk(j) - dtempk(j - 1))
                Next j
            Call surfint(ipoints, dtempk(), dy(), dalpha(), eroare)
            If eroare Then Err.Raise 1101, , "An errror encountered when trying to integrate the DTG curve."
            For i = 2 To ipoints
            ddalpha(i) = dalpha(i) / stotal#
            Next i
            ddalpha(1) = 0
            Case 3 ' tg, calculez alpha
            dalpha(1) = 0
            dalpha(ipoints) = 1
            For i = 2 To ipoints - 1
                dalpha(i) = (dy(i) - dy(1)) / (dy(ipoints) - dy(1))
            Next i
            Call deriv(ipoints, dalpha(), ddalpha(), eroare)
            If eroare Then Err.Raise 1101, , "Error trying to derivate the TG curve. Check your data."
        End Select
        txt = "    temp./C    temp./K  " & data_editor!lst2.Text & "  [alpha]     d[alpha]/dT "
        txt = txt & vbCrLf & CStr(1) & spatiu & Format$((dtempk(1) - 273.15), "#00.00") & spatiu & Format$(dtempk(1), "#000.00") & spatiu & Format$(dy(1), "####0.00###") & spatiu & Format$(dalpha(1), "0.0000")
       
        For i = 2 To ipoints - 1
        txt = txt & vbCrLf & CStr(i) & spatiu & Format$((dtempk(i) - 273.15), "#00.00") & spatiu & Format$(dtempk(i), "#000.00") & spatiu & Format$(dy(i), "####0.00###") & spatiu & Format$(dalpha(i), "0.0000") & spatiu & Format$(ddalpha(i), "0.00000")
        Next i
        txt = txt & vbCrLf & CStr(ipoints) & spatiu & Format$((dtempk(ipoints) - 273.15), "#00.00") & spatiu & Format$(dtempk(ipoints), "#000.00") & spatiu & Format$(dy(ipoints), "####0.00###") & spatiu & Format$(dalpha(ipoints), "0.0000")
        scrie_log txt
    If (data_editor!chkdif(0).Value = 1) Then   'face calculul primei metode dtg, adica Achar, sufix ac
        Call achar
    End If

    If (data_editor!chkdif(1).Value = 1) Then 'freeman caroll
        Call freeman_caroll
    End If
    
    If (data_editor!chkdif(2).Value = 1) Then
    'piloyan, sufix pi, determin doar energia de activare pentru domeniu restrans in alpha
    'intre 0.1 si la 0.5, am nevoie de cel putin 4 puncte aici altfel dau cu flitul
        Call piloyan
    End If

    If (data_editor!chkdif(3).Value = 1) Then
    'asta e a lu mitica, am datele din grid, trebuie sa fie t si dta
        Call fatu
    End If
    
Case "Regression"
'fac doar prin pseudo inversa
scrie_log "Pseudo-Iverse matrix method."
scrie_log "Conversion function : " & data_editor.lblconv.Caption
scrie_log "Heating rate, K/min: " & data_editor.txtreg.Text
scrie_log linie
    scrie_log linie & vbCrLf & "Attention: with regression you may obtain strange results." & vbCrLf & linie
    Dim ii() As Double, x() As Double, er As Boolean
    ReDim dalpha(ipoints), ddalpha(ipoints)
    Dim ind1 As Boolean, ind2 As Boolean, ind3 As Boolean 'indicatori asupra metodelor utilizate
    ind1 = CBool(data_editor!chkreg(0).Value) 'alpha^m
    ind2 = CBool(data_editor!chkreg(1).Value) '(1-alpha)^n
    ind3 = CBool(data_editor!chkreg(2).Value) '(-ln(1-alpha))^p
''    ReDim ind_regression(6) 'stochez toti parametri obtinuti in forma n,m,p, A,E, rate
        Select Case data_editor.lst2.ListIndex
        Case 0 ' alpha
            For i = 1 To ipoints: dalpha(i) = dy(i): Next i
            Call deriv(ipoints, dalpha(), ddalpha(), eroare)
            If eroare Then Err.Raise 1101, , "Error trying to derivate the conversion degree. Check your data."
        Case 1 ' dta,  treci la integrare
            Call surfint(ipoints, dtempk(), dy(), dalpha(), eroare)
            If eroare Then Err.Raise 1101, , "An errror encountered when trying to integrate the DTA curve."
            Call deriv(ipoints, dalpha(), ddalpha(), eroare)
            If eroare Then Err.Raise 1101, , "Error trying to derivate the DTA curve. Check your data."
        Case 2 ' dtg, integrare
             For j = 2 To ipoints
                stotal# = 0
                stotal# = stotal# + ((dy(j) + dy(j - 1)) / 2) * (dtempk(j) - dtempk(j - 1))
                Next j
            Call surfint(ipoints, dtempk(), dy(), dalpha(), eroare)
            If eroare Then Err.Raise 1101, , "An errror encountered when trying to integrate the DTG curve."
            For i = 2 To ipoints
            ddalpha(i) = dy(i) / stotal#
            Next i
            ddalpha(1) = 0
        Case 3 ' tg, calculez alpha
'            dalpha(1) = 0
 '           dalpha(ipoints) = 1
            For i = 1 To ipoints
                dalpha(i) = (dy(i) - dy(1)) / (dy(ipoints) - dy(1))
            Next i
            Call deriv(ipoints, dalpha(), ddalpha(), eroare)
            If eroare Then Err.Raise 1101, , "Error trying to derivate the TG curve. Check your data."
        End Select
        txt = "    temp./C    temp./K  " & data_editor!lst2.Text & "  [alpha]   d[alpha]/dT "
        txt = txt & vbCrLf & "1" & spatiu & Format$((dtempk(1) - 273.15), "#00.00") & spatiu & Format$(dtempk(1), "#000.00") & spatiu & Format$(dy(1), "####0.00###") & spatiu & Format$(dalpha(1), "0.00000")
        For i = 2 To ipoints - 1
        txt = txt & vbCrLf & CStr(i) & spatiu & Format$((dtempk(i) - 273.15), "#00.00") & spatiu & Format$(dtempk(i), "#000.00") & spatiu & Format$(dy(i), "####0.00###") & spatiu & Format$(dalpha(i), "0.0000") & spatiu & Format$(ddalpha(i), "#.0000000")
        Next i
        txt = txt & vbCrLf & CStr(ipoints) & spatiu & Format$((dtempk(ipoints) - 273.15), "#00.00") & spatiu & Format$(dtempk(ipoints), "#000.00") & spatiu & Format$(dy(ipoints), "####0.00###") & spatiu & Format$(dalpha(ipoints), "0.0000")
        scrie_log txt
 
    'optimizare, pseudoinversa
    ReDim zzz(ipoints - 2, 4), ii(ipoints - 2), x(4)
    'pun termenul liber si valorile pentru A si E, in coloana 1 si 2
    For i = 2 To ipoints - 1
    If ddalpha(i) = 0 Then ddalpha(i) = 0.0000000001
    ii(i - 1) = Log(rate / 60# * ddalpha(i))
    zzz(i - 1, 1) = 1
    zzz(i - 1, 2) = -1 / 8.31 / dtempk(i)
    Next i
    itest% = 0
        If ind1 Then
        'umple col. trei,pune in ind in primul element, numarul 3, arata coloana din matricea coeficientilor
            For i = 2 To ipoints - 1
            zzz(i - 1, 3) = Log(dalpha(i))
            Next i
            itest% = 1
        End If
        
        If ind2 Then
                For i = 2 To ipoints - 1
                zzz(i - 1, 3 + itest%) = Log(1 - dalpha(i))
                Next i
                itest% = itest% + 1
         End If

        If ind3 Then
                For i = 2 To ipoints - 1
                zzz(i - 1, 3 + itest%) = Log(-Log(1 - dalpha(i)))
                Next i
                itest% = itest% + 1
         End If

nec% = 2 + itest%
    
ReDim Preserve zzz(ipoints - 2, nec%), x(nec%)
'intregul nec% este numarul de necunoscute, este local
'calculez parametri cinetici prin pseudoinversa, asta trebuie mutata de aici ulterior
'am ipoints-2 ecuatii si 3 sau patru necunoscute
    Call pseudoinv(ipoints - 2, nec%, zzz(), ii(), x(), lowsize, eroare)
    If eroare Then Err.Raise 1101, , "Singular matrix."
    txt = "Pseudoinverse method "
    If x(2) < 1 Then txt = txt & vbCrLf & "Attention: check the conversion function !" & vbCrLf & "Activation energy ??"
    txt = txt & vbCrLf & linie
    txt = txt & vbCrLf & "Preexponential factor: " & Format$(Exp(x(1)), "+0.00##E+00") & " <1/sec.> "
    txt = txt & vbCrLf & "Activation energy : " & Format$(x(2), "#####0.0#") & " <J/mol> "
    txt = txt & vbCrLf & vbCrLf & "Conversion function: " & data_editor!lblconv.Caption
    scrie_log txt
    'fac testul de termen liber in txt2$
    txt2$ = "": rest# = 0: rest2# = 0
        
    Select Case itest%
        Case 1
            If ind1 Then
            txt = "    m is =" & Format$(x(3), "0.0##")
            txt2$ = " rate-A exp (-E/RT) f(alpha)" & Chr$(9) & "rate" & Chr$(9) & "Deviation %"
            For ij% = 2 To ipoints - 1
            rest# = rest# + CStr(ddalpha(ij%) - Exp(x(1)) * 60 / rate * Exp(-x(2) / 8.314 / dtempk(ij%)) * dalpha(ij%) ^ x(3))
            txt2$ = txt2$ & vbCrLf & Chr$(9) & Format$(rest#, "0.0000E+00") & Chr$(9) & Chr$(9) & Format$(ddalpha(ij%), "0.0000E+00") & Chr$(9) & Format$(Abs((rest#) / (ddalpha(ij%))) * 100, "####0.00")
            rest2# = rest2# + Abs((rest#) / (ddalpha(ij%))) * 100
            Next ij%
            txt2$ = txt2$ & vbCrLf & "Global deviation : " & Format$(rest2#, "######0.00")
            End If
            
            If ind2 Then
            txt = "    n is =" & Format$(x(3), "0.0##")
            txt2$ = " rate-A exp (-E/RT) f(alpha)" & Chr$(9) & "rate" & Chr$(9) & "Deviation %"
            For ij% = 2 To ipoints - 1
            rest# = (ddalpha(ij%) - Exp(x(1)) * 60 / rate * Exp(-x(2) / 8.314 / dtempk(ij%)) * (1 - dalpha(ij%)) ^ x(3))
            txt2$ = txt2$ & vbCrLf & Chr$(9) & Format$(rest#, "0.0000E+00") & Chr$(9) & Chr$(9) & Format$(ddalpha(ij%), "0.0000E+00") & Chr$(9) & Format$(Abs((rest#) / (ddalpha(ij%))) * 100, "####0.00")
            rest2# = rest2# + Abs((rest#) / (ddalpha(ij%))) * 100

            Next ij%
            txt2$ = txt2$ & vbCrLf & "Global deviation : " & Format$(rest2#, "######0.00")
            
            End If
            
            If ind3 Then
            txt = "    p is =" & Format$(x(3), "0.0##")
            txt2$ = " rate-A exp (-E/RT) f(alpha)" & Chr$(9) & "rate" & Chr$(9) & "Deviation %"
            For ij% = 2 To ipoints - 1
            rest# = (ddalpha(ij%) - Exp(x(1)) * 60 / rate * Exp(-x(2) / 8.314 / dtempk(ij%)) * (-Log(1 - dalpha(ij%))) ^ x(3))
            txt2$ = txt2$ & vbCrLf & Chr$(9) & Format$(rest#, "0.0000E+00") & Chr$(9) & Chr$(9) & Format$(ddalpha(ij%), "0.0000E+00") & Chr$(9) & Format$(Abs((rest#) / (ddalpha(ij%))) * 100, "####0.00")
            rest2# = rest2# + Abs((rest#) / (ddalpha(ij%))) * 100
           
            Next ij%
            txt2$ = txt2$ & vbCrLf & "Global deviation : " & Format$(rest2#, "######0.00")
            
            
            End If
            
            Case 2
            If ind1 Then
                    txt = "     m is =" & Format$(x(3), "0.0##")
'                    ind_regression(2) = X(3)
                    If ind2 Then
                        'm si n in col 3 si 4
                         txt = txt & vbCrLf & "     n is =" & Format$(x(4), "0.0##")
'                         ind_regression(1) = X(4) 'n
            txt2$ = " rate-A exp (-E/RT) f(alpha)" & Chr$(9) & "rate" & Chr$(9) & "Deviation %"
            For ij% = 2 To ipoints - 1
            rest# = (ddalpha(ij%) - Exp(x(1)) * 60 / rate * Exp(-x(2) / 8.314 / dtempk(ij%)) * ((1 - dalpha(ij%)) ^ x(4) * dalpha(ij%) ^ x(3)))
            txt2$ = txt2$ & vbCrLf & Chr$(9) & Format$(rest#, "0.0000E+00") & Chr$(9) & Chr$(9) & Format$(ddalpha(ij%), "0.0000E+00") & Chr$(9) & Format$((rest#) / (ddalpha(ij%)) * 100, "####0.00")
            rest2# = rest2# + Abs((rest#) / (ddalpha(ij%))) * 100
           
            Next ij%
            txt2$ = txt2$ & vbCrLf & "Global deviation : " & Format$(rest2#, "######0.00")
            
            
                     
                     Else
                 txt = txt & vbCrLf & "     p is =" & Format$(x(4), "0.0##")
'                 ind_regression(3) = X(4)
            txt2$ = " rate-A exp (-E/RT) f(alpha)" & Chr$(9) & "rate" & Chr$(9) & "Deviation %"
            For ij% = 2 To ipoints - 1
            rest# = (ddalpha(ij%) - Exp(x(1)) * 60 / rate * Exp(-x(2) / 8.314 / dtempk(ij%)) * ((-Log(1 - dalpha(ij%))) ^ x(4) * dalpha(ij%) ^ x(3)))
            txt2$ = txt2$ & vbCrLf & Chr$(9) & Format$(rest#, "0.0000E+00") & Chr$(9) & Chr$(9) & Format$(ddalpha(ij%), "0.0000E+00") & Chr$(9) & Format$((rest#) / (ddalpha(ij%)) * 100, "####0.00")
            rest2# = rest2# + Abs((rest#) / (ddalpha(ij%))) * 100
           
            Next ij%
            txt2$ = txt2$ & vbCrLf & "Global deviation : " & Format$(rest2#, "######0.00")
                 
                 
                 End If
            Else
                txt = "     n is =" & Format$(x(3), "0.0##")
                txt = txt & vbCrLf & "     p is =" & Format$(x(4), "0.0##")
'                ind_regression(1) = X(3) 'n
'                ind_regression(3) = X(4) 'p
            txt2$ = " rate-A exp (-E/RT) f(alpha)" & Chr$(9) & "rate" & Chr$(9) & "Deviation %"
            For ij% = 2 To ipoints - 1
            rest# = (ddalpha(ij%) - Exp(x(1)) * 60 / rate * Exp(-x(2) / 8.314 / dtempk(ij%)) * ((1 - dalpha(ij%)) ^ x(3) * (-Log(1 - dalpha(ij%)))) ^ x(4))
            txt2$ = txt2$ & vbCrLf & Chr$(9) & Format$(rest#, "0.0000E+00") & Chr$(9) & Chr$(9) & Format$(ddalpha(ij%), "0.0000E+00") & Chr$(9) & Format$((rest#) / (ddalpha(ij%)) * 100, "####0.00")
            rest2# = rest2# + Abs((rest#) / (ddalpha(ij%))) * 100
           
            Next ij%
            txt2$ = txt2$ & vbCrLf & "Global deviation : " & Format$(rest2#, "######0.00")
            
            
            End If
        Case 3
        Err.Raise 1101, , "Conversion function not accepted."
        End Select
scrie_log txt
scrie_log txt2$

gindicator(3) = False

Case "CRTA"
ReDim dalpha(ipoints1 + ipoints2 + 1)
'treci la crta, numai dupa ce ai facut toate verificarile
'If data_editor.optcrt(1).Value Then
'calculele mici legate de icar2
ReDim dalpha(ipoints1 + ipoints2 + 1)
txt = linie
txt = txt & vbCrLf & "ICAR 2 :  IsoKinetic - Analysis and Regression "
        For i = 1 To ipoints1: dalpha(i) = (dy(i) - ivm1) / (ivm2 - ivm1): Next i
        For i = ipoints1 + 2 To ipoints1 + ipoints2 + 1
            dalpha(i) = (dy(i) - iivm1) / (iivm2 - iivm1)
        Next i
'        txt=txt  &  ("    temp./C    temp./K  " & data_editor!lst2.Text & "  [alpha]  ")
'        For i = 1 To ipoints1
'            txt=txt  &  (CStr(i) & spatiu & Format$((dtempk(i) - 273.15), "###0.00") & spatiu & Format$(dtempk(i), "#000.00") & spatiu & Format$(dy(i), "####0.00###")) & spatiu & Format$(dalpha(i), "0.0000")
'        Next i
'    txt=txt  &  "Second data set:"
'        For i = ipoints1 + 2 To ipoints1 + ipoints2 + 1
'            txt=txt  &  (CStr(i - ipoints1 - 1) & spatiu & Format$((dtempk(i) - 273.15), "###0.00") & spatiu & Format$(dtempk(i), "#000.00") & spatiu & Format$(dy(i), "####0.00###")) & spatiu & Format$(dalpha(i), "0.0000")
'        Next i
scrie_log txt
    Call icar2
Case Else
    Err.Raise 1101, , "Unexpected error. Please report Error in menu_comp_param, case else methods type"
End Select
Beep
main_display.MousePointer = 0
main_display.richtxtlog.Visible = main_display.mnu_viewlog.Checked
Exit Sub
localhandle:
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
main_display.MousePointer = 0
main_display.richtxtlog.Visible = main_display.mnu_viewlog.Checked
Close
Exit Sub
End Sub


Sub urbanovici()
On Error GoTo handleit
Dim si As Double, fdealpha As Double, i As Integer, itest As Integer, panta As Double, ordonata As Double, corrs As Double, deltab As Double, deltam As Double, eroare As Boolean
ReDim linx(ipoints) As Double, liny(ipoints) As Double 'variabila locala
ReDim linx(ipoints) As Double, liny(ipoints) As Double
Dim aus As Double, eus As Double, mus As Double, bus As Double, dmus As Double, dbus As Double, cus As Double, nus As Double
    cus = 0
    Dim ius() As Double
    ReDim ius(ipoints) As Double
    scrie_log (linie & vbCrLf & "Computing Urbanovici-Segal values. ")
    scrie_log ("reaction order, slope, intercept, correlation")
    For si = ninitial To nfinal Step nstep
        For itest = 2 To ipoints - 1
     ius(itest) = ius(itest - 1) + (dalpha(itest) - dalpha(itest - 1)) / 2 * (1 / (1 - dalpha(itest)) ^ si / dtempk(itest) ^ 2 + 1 / (1 - dalpha(itest - 1)) ^ si / dtempk(itest - 1) ^ 2)
     liny(itest) = Log(ius(itest))
     linx(itest) = 1 / dtempk(itest)
        Next itest 'chem regresia
    Call reglin(ipoints, linx(), liny(), sigy, panta, ordonata, corrs, deltam, deltab, eroare)
        If eroare Then Err.Raise 1101, , "Error in the regression routine. Check your data."
    mesaj = Format$(si, "0.000") & spatiu & Format$(panta, "0.0000E+00") & spatiu & Format$(ordonata, "0.0000E+00") & spatiu & Format$(corrs, "0.00000000")
    If Abs(corrs) > Abs(cus) Then
       'am gasit un parametru mai bun, fac schimbarile
        mesaj = mesaj + ", it's better..."
        nus = si: mus = panta: bus = ordonata: dmus = deltam: dbus = deltab: cus = corrs
    End If
    scrie_log (mesaj)
    Next si
    'factorul preexp este acr
    eus = -mus * 8.314 'lucrez in jouli
    aus = Exp(bus) * rate / 60 * eus / 8.314
    deus = dmus * 8.314
    daus = Abs(aus - Exp(dbus + bus) * rate / 60 * eus / 8.314)
    scrie_log (linie & vbCrLf & "Parameters obtained for Urbanovici Segal method:")
    scrie_log ("Activation energy:  " & Format$(eus, "#####.##") & "  <J/mol>")
    scrie_log ("Activation energy, std. deviation:  " & Format$(deus, "######.#") & "  <J/mol>")
    scrie_log ("Preexponential factor:  " & Format$(aus, "0.00#E+00") & "  <1/sec>")
    scrie_log ("Preexponential factor, std. deviation:  " & Format$(daus, "0.0#E+00") & "  <1/sec>")
    scrie_log ("Correlation coefficient:  " & Format$(cus, "0.00000#####"))
    scrie_log ("Reaction order retained:  " & Format$(nus, "0.000##") & vbCrLf & linie)
gindicator(6) = True
Dim x(200) As Double, y(200) As Double
Call simulate(y(), x(), dtempk(1) - 1, dtempk(ipoints) + 1, nus, aus, eus, CDbl(rate / 60), 4, eroare)
If eroare Then
    gindicator(6) = False
Else
    For i% = 1 To 200
    xgraf(i%, 6) = x(i%): ygraf(i%, 6) = y(i%)
    Next i%
'ygraf(1, 6) = 0
End If
Exit Sub
handleit:
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
Close
Exit Sub
End Sub

Sub flynn_wall()
Dim si As Double, fdealpha As Double, txt As String
ReDim linx(ipoints) As Double, liny(ipoints) As Double 'variabila locala
On Error GoTo handleit
Dim i As Integer, itest As Integer, panta As Double, ordonata As Double, corrs As Double, deltab As Double, deltam As Double
Dim eroare As Boolean
Dim afw As Double, efw As Double, mfw As Double, bfw As Double
    Dim dmfw As Double, dbfw As Double, cfw As Double
    Dim nfw As Double
    cfw = 0
txt = linie & vbCrLf & "Computing Flynn Wall values. "
    txt = txt & vbCrLf & "reaction order, slope, intercept, correlation"
    For si = ninitial To nfinal Step nstep
        For itest = 2 To ipoints - 1
        If (si < 1.0001) And (si > 0.9999) Then
          fdealpha = -Log(1 - dalpha(itest))
        Else
          fdealpha = (1 - (1 - dalpha(itest)) ^ (1 - si)) / (1 - si)
        End If
            liny(itest) = Log(fdealpha) ' / Log(10#)
            linx(itest) = 1 / dtempk(itest)
        Next itest
    Call reglin(ipoints, linx(), liny(), sigy, panta, ordonata, corrs, deltam, deltab, eroare)
        If eroare Then Err.Raise 1101, , "Error in the regression routine (Flynn Wall). Check your data."
    txt = txt & vbCrLf & Format$(si, "0.000") & spatiu & Format$(panta, "0.0000E+00") & spatiu & Format$(ordonata, "0.0000E+00") & spatiu & Format$(corrs, "0.00000000")
        If Abs(corrs) > Abs(cfw) Then
       'am gasit un parametru mai bun, fac schimbarile
    txt = txt & ", it's better..."
        nfw = si: mfw = panta: bfw = ordonata: dmfw = deltam: dbfw = deltab: cfw = corrs
        End If
        Next si
    scrie_log txt
    efw = -mfw * 8.314 / 1.052 'lucrez in jouli
    afw = 8.314 * (rate / 60) / efw * Exp(bfw + 5.33)
    defw = dmfw * 8.314 / 1.052
    dafw = Abs(afw - 8.134 * (rate / 60) / (-defw + efw) * Exp(bfw + dbfw + 5.33))
    
'    dafw = 1.986 / 60 * rate * 10 ^ (dbfw + 5.33) / efw * 1.986 / 8.31
txt = linie & vbCrLf & "Parameters obtained for Flynn Wall method:"
txt = txt & vbCrLf & "Activation energy:  " & Format$(efw, "#####.0#") & "  <J/mol>"
txt = txt & vbCrLf & "Activation energy, std. deviation:  " & Format$(defw, "######.0") & "  <J/mol>"
txt = txt & vbCrLf & "Preexponential factor:  " & Format$(afw, "0.000#E+00") & "  <1/sec>"
txt = txt & vbCrLf & "Preexponential factor, std. deviation:  " & Format$(dafw, "0.0#E+00") & "  <1/sec>"
txt = txt & vbCrLf & "Correlation coefficient:  " & Format$(cfw, "0.00000#####")
scrie_log txt & vbCrLf & "Reaction order retained:  " & Format$(nfw, "0.000##") & vbCrLf & linie

gindicator(4) = True
Dim x(200) As Double, y(200) As Double
Call simulate(y(), x(), dtempk(1) - 1, dtempk(ipoints) + 1, nfw, afw, CDbl(efw), rate / 60#, 4, eroare)
If eroare Then
    gindicator(4) = False
Else
    For i% = 1 To 200
    xgraf(i%, 4) = x(i%): ygraf(i%, 4) = y(i%)
    Next i%
End If
Exit Sub
handleit:
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
Close
Exit Sub
End Sub
Sub piloyan()
On Error GoTo handleit
Dim si As Double, fdealpha As Double
ReDim linx(ipoints) As Double, liny(ipoints) As Double 'variabila locala
Dim i As Integer, itest As Integer, panta As Double, ordonata As Double, corrs As Double, deltab As Double, deltam As Double, eroare As Boolean
  
    Dim epi As Double, mpi As Double
    Dim dmpi As Double, dbpi As Double, cpi As Double
    cpi = 0
    scrie_log (linie & vbCrLf & "Computing Piloyan values. ")
    scrie_log ("Attention: special conditions requested for this method")
        For itest = 2 To ipoints - 1
    If dalpha(itest) > 0.6 Then
    pil% = itest
    If pil% < 3 Then Err.Raise 1101, , "I need more date for alpha<0.5 in the Piloyan method."
    Exit For
    End If
            liny(itest) = Log(ddalpha(itest))
            linx(itest) = 1 / dtempk(itest)
    Next itest
    Call reglin(pil%, linx(), liny(), sigy, panta, ordonata, corrs, deltam, deltab, eroare)
        If eroare Then Err.Raise 1101, , "Error in the regression routine (Piloyan). Check your data."
      epi = -panta * 8.314 'lucrez in jouli
    depi = deltam * 8.314
    scrie_log (linie & vbCrLf & "Parameters obtained for Piloyan method:")
    scrie_log ("Activation energy (computed for 0.1 < alpha < 0.6 :  " & Format$(epi, "#####.##") & "  <J/mol>")
    scrie_log ("Activation energy, std. deviation:  " & Format$(depi, "######.#") & "  <J/mol>")
    scrie_log ("Correlation coefficient:  " & Format$(corrs, "0.00000#####"))
gindicator(5) = False 'intotdeauna
Exit Sub
handleit:
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
Close
Exit Sub
End Sub
Sub van_krevelen()
On Error GoTo handleit
Dim i As Integer, itest As Integer, panta As Double, ordonata As Double, corrs As Double, deltab As Double, deltam As Double
Dim eroare As Boolean
ReDim linx(ipoints) As Double, liny(ipoints) As Double
    Dim avk As Double, evk As Double, mvk As Double, bvk As Double
    Dim dmvk As Double, dbvk As Double, cvk As Double
    Dim nvk As Double
    cvk = 0
    scrie_log (linie & vbCrLf & "Computing van Krevelen values. " & vbCrLf & "Attention: there are special restriction in this method. ")
    scrie_log ("You may obtain some strange results.")
        scrie_log ("reaction order, slope, intercept, correlation")
    For si = ninitial To nfinal Step nstep
        For itest = 2 To ipoints - 1
        If (si < 1.00001) And (si > 0.999999) Then
          fdealpha = -Log(1 - dalpha(itest))
        Else
          fdealpha = (1 - (1 - dalpha(itest)) ^ (1 - si)) / (1 - si)
        End If
            liny(itest) = Log(fdealpha) / Log(10#)
            linx(itest) = Log(dtempk(itest)) / Log(10#)
        Next itest 'chem regresia
    Call reglin(ipoints, linx(), liny(), sigy, panta, ordonata, corrs, deltam, deltab, eroare)
        If eroare Then Err.Raise 1101, , "Error in the regression routine (van Krevelen). Check your data."
    mesaj = Format$(si, "0.000") & spatiu & Format$(panta, "0.0000E+00") & spatiu & Format$(ordonata, "0.0000E+00") & spatiu & Format$(corrs, "0.00000000")
    If Abs(corrs) > Abs(cvk) Then
       'am gasit un parametru mai bun, fac schimbarile
        mesaj = mesaj + ", it's better..."
        nvk = si: mvk = panta: bvk = ordonata: dmvk = deltam: dbvk = deltab: cvk = corrs
    End If
    scrie_log (mesaj)
    Next si
    evk = -8.314 * dtempk(ipoints) * (1 - mvk)
    avk = rate / 60 * 10 ^ bvk * 1 / ((0.368 / dtempk(ipoints)) ^ (evk / 8.314 / dtempk(ipoints)) * ((8.314 * dtempk(ipoints)) / (evk + 8.314 * dtempk(ipoints))))
    devk = Abs(evk + 8.314 * dtempk(ipoints) * (1 - (dmvk + mvk)))
    davk = Abs(avk - rate / 60 * 10 ^ (dbvk + bvk) * 1 / ((0.368 / dtempk(ipoints)) ^ (evk / 8.314 / dtempk(ipoints)) * ((8.314 * dtempk(ipoints)) / (evk + 8.314 * dtempk(ipoints)))))
    scrie_log (linie & vbCrLf & "Parameters obtained for van Krevelen method:")
    scrie_log ("Activation energy:  " & Format$(evk, "#####.##") & "  <J/mol>")
    scrie_log ("Activation energy, std. deviation:  " & Format$(devk, "######.#") & "  <J/mol>")
    scrie_log ("Preexponential factor:  " & Format$(avk, "0.000#E+00") & "  <1/sec>")
    scrie_log ("Preexponential factor, std. deviation:  " & Format$(davk, "0.0#E+00") & "  <1/sec>")
    scrie_log ("Correlation coefficient:  " & Format$(cvk, "0.00000#####"))
    scrie_log ("Reaction order retained:  " & Format$(nvk, "0.000##") & vbCrLf & linie)
gindicator(5) = True
Dim x(200) As Double, y(200) As Double
Call simulate(y(), x(), dtempk(1) - 1, dtempk(ipoints) + 1, nvk, avk, evk, CDbl(rate / 60), 4, eroare)
If eroare Then
    gindicator(5) = False
Else
    For i% = 1 To 200
    xgraf(i%, 5) = x(i%): ygraf(i%, 5) = y(i%)
    Next i%
 '   ygraf(1, 5) = 0
End If
Exit Sub
handleit:
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
Close
Exit Sub
End Sub
Sub achar()
On Error GoTo handleit
Dim aac As Double, eac As Double, mac As Double, bac As Double
Dim dmac As Double, dbac As Double, cac As Double
Dim nac As Double, txt As String
ReDim linx(ipoints) As Double, liny(ipoints) As Double
Dim si As Double, fdealpha As Double
Dim i As Integer, itest As Integer, panta As Double, ordonata As Double, corrs As Double, deltab As Double, deltam As Double, eroare As Boolean
    cac = 0
    txt = (linie & vbCrLf & "Computing Achar values. ")
    scrie_log txt & vbCrLf & "reaction order,  slope,  intercept,  correlation"
    For si = ninitial To nfinal Step nstep
        For itest = 2 To ipoints - 1
            liny(itest) = Log((ddalpha(itest)) / ((1 - dalpha(itest))) ^ si)
            linx(itest) = 1 / dtempk(itest)
        Next itest
    Call reglin(ipoints, linx(), liny(), sigy, panta, ordonata, corrs, deltam, deltab, eroare)
        If eroare Then Err.Raise 1101, , "Error in the regression routine (Achar). Check your data."
    txt = txt & vbCrLf & Format$(si, "0.000") & spatiu & Format$(panta, "0.0000E+00") & spatiu & Format$(ordonata, "0.0000E+00") & spatiu & Format$(corrs, "0.00000000")
    If Abs(corrs) > Abs(cac) Then
       'am gasit un parametru mai bun, fac schimbarile
        txt = txt & ", it's better..."
        nac = si: mac = panta: bac = ordonata: dmac = deltam: dbac = deltab: cac = corrs
    End If
    Next si
    eac = -mac * 8.314 'lucrez in jouli
    aac = rate / 60 * Exp(bac)
    deac = dmac * 8.314
    daac = Abs(aac - (rate / 60) * (Exp(bac + dbac)))
    txt = txt & vbCrLf & linie & vbCrLf & "Parameters obtained for Achar method:"
    txt = txt & vbCrLf & "Activation energy:  " & Format$(eac, "#####.##") & "  <J/mol>"
    txt = txt & vbCrLf & "Activation energy, std. deviation:  " & Format$(deac, "######.#") & "  <J/mol>"
    txt = txt & vbCrLf & "Preexponential factor:  " & Format$(aac, "0.000#E+00") & "  <1/sec>"
    txt = txt & vbCrLf & "Preexponential factor, std. deviation:  " & Format$(daac, "0.0#E+00") & "  <1/sec>"
    txt = txt & vbCrLf & "Correlation coefficient:  " & Format$(cac, "0.00000#####")
    txt = txt & vbCrLf & "Reaction order:  " & Format$(nac, "0.000##") & vbCrLf & linie
    scrie_log txt
gindicator(3) = True
Dim x(200) As Double, y(200) As Double
Call simulate(y(), x(), dtempk(1) - 1, dtempk(ipoints) + 1, nac, aac, eac, CDbl(rate / 60), 5, eroare)
If eroare Then gindicator(3) = False
For i% = 1 To 200
xgraf(i%, 3) = x(i%): ygraf(i%, 3) = y(i%)
Next i%
Exit Sub
handleit:
scrie_log txt
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
Close
Exit Sub
End Sub
Sub fatu()
On Error GoTo handleit
gindicator(6) = False
Dim adf As Double, edf As Double, mdf As Double, bdf As Double, txt As String
    Dim dbdf As Double, dmdf As Double, cdf As Double
    Dim ndf As Double
Dim si As Double, fdealpha As Double
Dim i As Integer, itest As Integer, panta As Double, ordonata As Double, corrs As Double, deltab As Double, deltam As Double, eroare As Boolean
ReDim linx(ipoints) As Double, liny(ipoints) As Double
    cdf = 0 'coef lui d fatu
    txt = linie & vbCrLf & "DTA data, D. Fatu's procedure."
  
         stotal# = 0
                For j% = 2 To ipoints
                stotal# = stotal# + ((dy(j%) + dy(j% - 1)) / 2) * (dtempk(j%) - dtempk(j% - 1))
                Next j%
        scrie_log txt & vbCrLf & "Total surface :" & "  " & Format$(stotal#, "######0.##########")
    txt = ""
        For si = ninitial + nstep To nfinal Step nstep
    For itest = 2 To ipoints - 1
    linx(itest) = 1 / dtempk(itest)
   liny(itest) = Log(dy(itest) / (1 - dalpha(itest)) ^ si)
    Next itest
    Call reglin(ipoints, linx(), liny(), sigy, panta, ordonata, corrs, deltam, deltab, eroare)
    If eroare Then Err.Raise 1101, , "Error in the regression routine. Check your data."
    txt = txt & vbCrLf & Format$(si, "0.0000") & spatiu & Format$(panta, "0.00000E+00") & spatiu & Format$(ordonata, "0.00000E+00") & spatiu & Format$(corrs, "0.00000000")
    If Abs(corrs) > Abs(cdf) Then
       'am gasit un parametru mai bun, fac schimbarile
        txt = txt + ",  better correlation..."
        ndf = si: mdf = panta: bdf = ordonata: dmdf = deltam: dbdf = deltab: cdf = corrs
    End If
    Next si
scrie_log txt
    edf = -mdf * 8.314 'lucrez in jouli
    adf = Exp(bdf) / stotal# / 60
    dedf = dmdf * 8.314
    dadf = Abs(adf - Exp(bdf + dbdf) / stotal# / 60)
txt = linie & vbCrLf & "Parameters obtained for Fatu method:"
txt = txt & vbCrLf & "Activation energy:  " & Format$(edf, "#####.##") & "  <J/mol>"
txt = txt & vbCrLf & "Activation energy, std. deviation:  " & Format$(dedf, "######.#") & "  <J/mol>"
txt = txt & vbCrLf & "Preexponential factor:  " & Format$(adf, "0.00##E+00") & "  <1/sec>"
txt = txt & vbCrLf & "Preexponential factor, std. deviation:  " & Format$(dadf, "0.0##E+00") & "  <1/sec>"
txt = txt & vbCrLf & "Correlation coefficient:  " & Format$(cdf, "0.00000####")
txt = txt & vbCrLf & "Reaction order retained:  " & Format$(ndf, "0.000##") & vbCrLf & linie
'dif_indicator(4, 1) = ndf: dif_indicator(4, 2) = adf: dif_indicator(4, 3) = edf: dif_indicator(4, 4) = rate
scrie_log txt
gindicator(6) = True
Dim x(200) As Double, y(200) As Double
Call simulate(y(), x(), dtempk(1) - 1, dtempk(ipoints) + 1, ndf, adf, edf, CDbl(rate / 60), 5, eroare)
If eroare Then gindicator(6) = False
For i% = 1 To 200
xgraf(i%, 6) = x(i%): ygraf(i%, 6) = y(i%)
Next i%

Exit Sub
handleit:
scrie_log txt
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
Close
Exit Sub
End Sub
Sub freeman_caroll()
On Error GoTo handleit
gindicator(4) = False
Dim txt As String
txt = vbCrLf & "Freeman Caroll method."
    Dim fcx As Double, fcy As Double, s1 As Double, s2 As Double, s3 As Double, s4 As Double, s5 As Double, s6 As Double
    For itest = 2 To ipoints - 1
    For j% = itest + 1 To ipoints - 1
    fcx = (1 / dtempk(j%) - 1 / dtempk(itest)) / Log(((1 - dalpha(j%)) / (1 - dalpha(itest))))
    fcy = (Log(ddalpha(j%) / ddalpha(itest))) / (Log((1 - dalpha(j%)) / (1 - dalpha(itest))))
    s1 = s1 + fcx: s2 = s2 + fcx * fcx: s3 = s3 + fcy: s4 = s4 + fcy * fcy: s5 = s5 + fcx * fcy
    s6 = s6 + 1
    Next j%
    Next itest
    panta = (s5 - s1 * s3 / s6) / (s2 - s1 * s1 / s6)
    deltam = Sqr(Abs((sigy * sigy) / (s6 * (s2 - s1 * s1))))
    ordonata = s3 / 56 - panta * s1 / s6
    deltab = Sqr(Abs((sigy * sigy * s2) / (s6 * (s2 - s1 * s1))))
    corrs = (s5 - s1 * s3 / s6) / Sqr(Abs((s2 - s1 * s1 / s6) * (s4 - s3 * s3 / s6)))
    Dim z, dz As Double
    z = 0
    For itest = 2 To ipoints - 1
    z = z + rate / 60 * ddalpha(itest) * Exp(-panta / dtempk(itest)) / ((1 - dalpha(itest)) ^ ordonata)
    Next itest
    z = z / (ipoints - 2)
    For itest = 2 To ipoints - 1
    dz = dz + (z - rate / 60 * ddalpha(itest) * Exp(-panta / dtempk(itest)) / ((1 - dalpha(itest)) ^ ordonata)) ^ 2
    Next itest
    dz = Sqr(dz / (ipoints - 2))
    
    txt = txt & vbCrLf & linie & vbCrLf & "Parameters obtained for Freeman Caroll method:"
    txt = txt & vbCrLf & "Attention: reaction order may be out of the specified domain."
    txt = txt & vbCrLf & "Reaction order :  " & Format$(ordonata, "0.000##")
    txt = txt & vbCrLf & "Reaction order, std. dev. :  " & Format$(deltab, "0.000##")
    txt = txt & vbCrLf & "Activation energy:  " & Format$((-panta * 8.314), "#####.##") & "  <J/mol>"
    txt = txt & vbCrLf & "Activation energy, std. dev.:  " & Format$(deltam * 8.314, "#####.##") & "  <J/mol>"
    txt = txt & vbCrLf & "Preexponential factor:  " & Format$(z, "0.00##E+00") & "  <1/sec>"
    txt = txt & vbCrLf & "Preexponential factor, std. deviation:  " & Format$(dz, "0.0##E+00") & "  <1/sec>"
    txt = txt & vbCrLf & "Correlation coefficient:  " & Format$(corrs, "0.00000####")
scrie_log txt
gindicator(4) = True
Dim x(200) As Double, y(200) As Double
Call simulate(y(), x(), dtempk(1) - 1, dtempk(ipoints) + 1, CDbl(ordonata), CDbl(z), CDbl(-panta * 8.314), CDbl(rate / 60), 4, CBool(eroare))
If eroare Then gindicator(4) = False
For i% = 1 To 200
xgraf(i%, 4) = x(i%): ygraf(i%, 4) = y(i%)
Next i%

Exit Sub
handleit:
scrie_log txt
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
Close
Exit Sub
End Sub

Sub coats_redfern()
On Error GoTo handleit
Dim si As Double, fdealpha As Double, txt As String
ReDim linx(ipoints) As Double, liny(ipoints) As Double 'variabila locala
Dim i As Integer, itest As Integer, panta As Double, ordonata As Double, corrs As Double, deltab As Double, deltam As Double
Dim stest1 As Double, stest2 As Double
Dim eroare As Boolean
Dim acr As Double, ecr As Double, mcr As Double, bcr As Double ', sqchi As Double
Dim dmcr As Double, dbcr As Double, ccr As Double, ncr As Double, stest As Double
    ccr = 0
    scrie_log linie & vbCrLf & "Computing Coats Redfern values. " & vbCrLf & "reaction order, slope, intercept, correlation"
    txt = ""
'calculez sig pentru fiecare y, am eroarea in procente, pentru alpha
'folosesc expresia de compunere a erorilor din cartea de statistica a lui
'barlow, adica sigma F= (dF/dx)sigma x
'stest = 0
'   For itest = 2 To ipoints - 1
'   stest = stest + (1 / (Log(1 - dalpha(itest))) * 1 / (1 - dalpha(itest)) * 1 / (Log(-Log(1 - dalpha(itest)) / dtempk(itest) / dtempk(itest)))) * sigy
'   Next itest
'    stest = Abs(stest / 100 / (ipoints - 2))
    For si = ninitial To nfinal Step nstep
'stest1 = 0: stest2 = 0
        For itest = 2 To ipoints - 1
        If (si < 1.00001) And (si > 0.999999) Then
          fdealpha = -Log(1 - dalpha(itest))
'stest1 = stest1 + (1 / (1 - dalpha(itest))) * sigy / 100 * dalpha(itest)
          Else
          fdealpha = (1 - (1 - dalpha(itest)) ^ (1 - si)) / (1 - si)
'stest1 = stest1 + (1 / (fdealpha)) * (-(1 - si) * (1 - dalpha(itest)) ^ (-si)) / ((1 - si)) * sigy / 100 * dalpha(itest)
          End If
            liny(itest) = Log((fdealpha) / (dtempk(itest) * dtempk(itest)))
            linx(itest) = 1 / dtempk(itest)
        Next itest 'chem regresia
        Call reglin(ipoints, linx(), liny(), sigy, panta, ordonata, corrs, deltam, deltab, eroare)
        If eroare Then Err.Raise 1101, , "Error in the regression routine (Coats Redfern). Check your data."
    txt = txt & vbCrLf & Format$(si, "0.000") & spatiu & Format$(panta, "0.0000E+00") & spatiu & Format$(ordonata, "0.0000E+00") & spatiu & Format$(corrs, "0.00000000") '& spatiu & Format$(sqchi, "######0.00000000")
    If Abs(corrs) > Abs(ccr) Then
       'am gasit un parametru mai bun, fac schimbarile
        txt = txt & ", it's better..."
        ncr = si: mcr = panta: bcr = ordonata: dmcr = deltam: dbcr = deltab: ccr = corrs
    End If
        Next si
    scrie_log txt
'factorul preexp este acr
    ecr = -mcr * 8.314 'lucrez in jouli
   ''''atentie la formula
   acr = Exp(bcr) * rate / 60 * ecr / 8.314
    decr = dmcr * 8.314
    dacr = Abs(acr - Exp(dbcr + bcr) * rate / 60 * ecr / 8.314)
txt = ""
    txt = txt & vbCrLf & linie & vbCrLf & "Parameters obtained for Coats-Redfern method:"
txt = txt & vbCrLf & "Activation energy:  " & Format$(ecr, "#####.##") & "  <J/mol>"
txt = txt & vbCrLf & "Activation energy, std. deviation:  " & Format$(decr, "######.#") & "  <J/mol>"
txt = txt & vbCrLf & "Preexponential factor:  " & Format$(acr, "0.00#E+00") & "  <1/sec>"
txt = txt & vbCrLf & "Preexponential factor, std. deviation:  " & Format$(dacr, "0.0#E+00") & "  <1/sec>"
txt = txt & vbCrLf & "Correlation coefficient:  " & Format$(ccr, "0.00000#####")
scrie_log txt & vbCrLf & "Reaction order:  " & Format$(ncr, "0.000##") & vbCrLf & linie

gindicator(3) = True
' simulate(al() As Double, te() As Double, tempstart As Double, tempend As Double, n As Double, a As Double, e As Double, r As Double, npx As Integer, eroare As Boolean)
Dim x(200) As Double, y(200) As Double
Call simulate(y(), x(), dtempk(1) - 1, dtempk(ipoints) + 1, ncr, acr, CDbl(ecr), rate / 60#, 4, eroare)
If eroare Then
    gindicator(3) = False
Else
    For i% = 1 To 200
    xgraf(i%, 3) = x(i%): ygraf(i%, 3) = y(i%)
    Next i%
End If
Exit Sub
handleit:
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
gindicator(3) = False
Close
Exit Sub
End Sub


Sub icar2()
ReDim coef(11, 2)
For i = 1 To 7
gindicator(i) = False
Next i
On Error GoTo handleit
            'icar 2'--------------------------------
            'grpol este gradul polinomului pentru interpolare, excl este domeniul de excludere, double global
            'nrint 'in loc de nre, este numarul de intervale + 1, il citesc in data editor, global
Dim eroare As Boolean, txt As String, cc(2) As Double
ReDim mmval(nrint, 2), nnval(nrint, 2), zzz(ipoints1 + ipoints2, 11)
'ipoints1 si 2 sunt numarul de puncte, globale, intregi
cc(1) = rate1 / 60 / (ivm2 - ivm1)
cc(2) = rate2 / 60 / (iivm2 - iivm1) 'vitezele celor doua procese in sec-1
'calculez pentru ambele curbe ??
For i = 1 To ipoints1
 For k = 1 To grpol
 zzz(i, k) = dalpha(i) ^ (k - 1)
 Next k 'alpha e x
Next i
 'atentie, modificat temporar
Dim tt() As Double, zzz2() As Double, tint() As Double, en() As Double
Dim min(), max() As Double, ap() As Double
Dim tempmin(2), tempmax(2), derivata(2)
ReDim min(3), max(3), s(grpol)
ReDim tt(ipoints2), zzz2(ipoints2, grpol)
ReDim tint(nrint, 2), en(nrint), ap(nrint, 2)
Dim alphamin(2) As Double, nval(2) As Double, mval(2) As Double, amed(2) As Double
Call pseudoinv(ipoints1, grpol, zzz(), dtempk(), s(), lowsize, eroare)
If eroare Then Err.Raise 1101, , "Error in the inversion matrix routine."

For j% = 1 To grpol
coef(j%, 1) = s(j%)
Next j%
'in matricea coef pastrez coeficientii polinomului
For i = ipoints1 + 2 To ipoints1 + ipoints2 + 1
tt(i - ipoints1 - 1) = dtempk(i)
 For k = 1 To grpol
 zzz2(i - ipoints1 - 1, k) = dalpha(i) ^ (k - 1)
 Next k 'alpha e x
Next i

't este y, dat in kelvin
Call pseudoinv(ipoints2, grpol, zzz2(), tt(), s(), lowsize, eroare)
If eroare Then Err.Raise 1101, , "Error in the inversion matrix routine."
For j% = 1 To grpol
coef(j%, 2) = s(j%)
Next j%
'in matricea coef pastrez coeficientii polinomului

'indicele 3 este pentru intersectie
min(1) = 1
max(1) = 0
For i = 1 To ipoints1
If min(1) > dalpha(i) Then min(1) = dalpha(i)
If max(1) < dalpha(i) Then max(1) = dalpha(i)
Next i

'determin domeniul extrem pe axa x , deci pe alpha pentru a doua curba
'ar fi mers usor cu un SUB, lenesu' mai mult alearga
min(2) = 1
max(2) = 0
For i = ipoints1 + 2 To ipoints1 + ipoints2 + 1
If min(2) > dalpha(i) Then min(2) = dalpha(i)
If max(2) < dalpha(i) Then max(2) = dalpha(i)
Next i

min(3) = min(1)
If min(1) < min(2) Then min(3) = min(2)
max(3) = max(1)
If max(1) > max(2) Then max(3) = max(2)
If max(3) <= min(3) Then Err.Raise 1101, , "I can not compute, there are not common alpha values"
''vad daca in domeniul comun exista alpha=1/2
'energia de activare=-R*(delta lnC)/(delta(1/T))
rdlc = -8.31 * (Log(cc(1) / cc(2))) 'rezultatul in J/mol
'asta e pentru doua puncte, adaptabil
enmed = 0

For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
'en este un vector pana la nre
'nu e mare lucru cu DEF FN

For j% = 1 To 2
tint(i, j%) = igrec(coef(1, j%), coef(2, j%), coef(3, j%), coef(4, j%), coef(5, j%), coef(6, j%), coef(7, j%), coef(8, j%), coef(9, j%), coef(10, j%), coef(11, j%), r)
'MsgBox CStr(tint(i, j%) & " " & CStr(j%))
Next j%
'valorile interpolate ale temperaturii pentru cele doua curbe
en(i) = rdlc / (1 / tint(i, 1) - 1 / tint(i, 2))
enmed = enmed + en(i)
Next i
enmed = enmed / nrint
If (min(3) > 0.5 Or max(3) < 0.5) Then MsgBox "I can not compute A, m and n. I need 0.5 alpha value.": GoTo tiparire
'calculez abaterea standard a energiei de activare ; estdev
estddev# = 0
For i = 1 To nrint
estddev# = estddev# + (enmed - en(i)) * (enmed - en(i))
Next i
estddev# = Sqr(estddev# / nrint)
'calculul lui m si n
'

For kk = 1 To 2
minim = 9999# 'asta da temperatura minima !
maxim = 0#
For i = 1 To nrint
If tint(i, kk) < minim Then minim = tint(i, kk): alphamin(kk) = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If tint(i, kk) > maxim Then maxim = tint(i, kk)
Next i
tempmin(kk) = minim
tempmax(kk) = maxim
Next kk
If tempmin(1) < 100 Or tempmin(2) < 100 Then Err.Raise 1101, , "Error in the interpolation function. Check your data."
If tempmin(1) > 1500 Or tempmin(2) > 1500 Then Err.Raise 1101, , "Error in the interpolation function. Check your data."
''imi folosesc la grafic
'alphamin este valoarea lui alpha pentru care temperatura are un minim
'tempmin este valoarea acesteia - nu cred ca imi va folosi da' never know
'vad daca acest minim se afla la un capat (din domeniul comun):
'daca este in stanga atunci nval=0
'daca este in dreapta atunci mval=0
'daca e intermediar atunci ma folosesc de m/n=alphamin/(1-alphamin)
'celalalt il determin din ecuatia derivatei
'nval si mval sunt n si m din functia de conversie
'daca nu am alpha=1/2 am atentionat
r = 0.5
For kk = 1 To 2
derivata(kk) = (coef(2, kk) + 2 * coef(3, kk) * r + 3 * coef(4, kk) * r ^ 2 + 4 * coef(5, kk) * r ^ 3 + 5 * coef(6, kk) * r ^ 4 + 6 * coef(7, kk) * r ^ 5 + 7 * coef(8, kk) * r ^ 6 + 8 * coef(9, kk) * r ^ 7 + 9 * coef(10, kk) * r ^ 8 + 10 * coef(11, kk) * r ^ 9)
Next kk

Dim mstd(2) As Double, nstd(2) As Double, astd(2) As Double
For kk = 1 To 2
Select Case alphamin(kk)
Case max(3)
nval(kk) = 0
mvalmed = 0
exclud1 = 0 'un numarator care imi arata de cate ori am adaugat la medie
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(kk) - r) < (excl * r) Then exclud1 = exclud1 + 1: GoTo salt
mmval(i, kk) = -r * (coef(2, kk) + 2 * coef(3, kk) * r + 3 * coef(4, kk) * r ^ 2 + 4 * coef(5, kk) * r ^ 3 + 5 * coef(6, kk) * r ^ 4 + 6 * coef(7, kk) * r ^ 5 + 7 * coef(8, kk) * r ^ 6) * enmed / 8.31 / ((igrec(coef(1, kk), coef(2, kk), coef(3, kk), coef(4, kk), coef(5, kk), coef(6, kk), coef(7, kk), coef(8, kk), coef(9, kk), coef(10, kk), coef(11, kk), 0.5)) ^ 2)
mvalmed = mvalmed + mmval(i, kk)
salt:
Next i
mval(kk) = mvalmed / (nrint - exclud1)
exclud1 = 0 'un numarator care imi arata de cate ori am adaugat la medie

For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(kk) - r) < (excl * r) Then exclud1 = exclud1 + 1: GoTo saltstd1
mmval(i, kk) = -r * (coef(2, kk) + 2 * coef(3, kk) * r + 3 * coef(4, kk) * r ^ 2 + 4 * coef(5, kk) * r ^ 3 + 5 * coef(6, kk) * r ^ 4 + 6 * coef(7, kk) * r ^ 5 + 7 * coef(8, kk) * r ^ 6) * enmed / 8.31 / ((igrec(coef(1, kk), coef(2, kk), coef(3, kk), coef(4, kk), coef(5, kk), coef(6, kk), coef(7, kk), coef(8, kk), coef(9, kk), coef(10, kk), coef(11, kk), 0.5)) ^ 2)
mstd(kk) = mstd(kk) + (mmval(i, kk) - mvalmed) * (mmval(i, kk) - mvalmed)
saltstd1:
Next i
mstd(kk) = Sqr(mstd(kk) / (nrint - exclud1))



Case min(3)
mval(kk) = 0
nvalmed = 0
exclud2 = 0 'un numarator care imi arata de cate ori am adaugat la medie
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(kk) - r) < (excl * r) Then exclud2 = exclud2 + 1: GoTo salt2
nnval(i, kk) = (1 - r) * (coef(2, kk) + 2 * coef(3, kk) * r + 3 * coef(4, kk) * r ^ 2 + 4 * coef(5, kk) * r ^ 3 + 5 * coef(6, kk) * r ^ 4 + 6 * coef(7, kk) * r ^ 5 + 7 * coef(8, kk) * r ^ 6 + 8 * coef(9, kk) * r ^ 7 + 9 * coef(10, kk) * r ^ 8 + 10 * coef(11, kk) * r ^ 9) * enmed / 8.31 / ((igrec(coef(1, kk), coef(2, kk), coef(3, kk), coef(4, kk), coef(5, kk), coef(6, kk), coef(7, kk), coef(8, kk), coef(9, kk), coef(10, kk), coef(11, kk), 0.5)) ^ 2)
nvalmed = nvalmed + nnval(i, kk)
salt2:
Next i
nval(kk) = nvalmed / (nrint - exclud2) '

exclud2 = 0 'un numarator care imi arata de cate ori am adaugat la medie
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(kk) - r) < (excl * r) Then exclud2 = exclud2 + 1: GoTo saltstd2
nstd(kk) = nstd(kk) + (nnval(i, kk) - nval(kk)) * (nnval(i, kk) - nval(kk))
saltstd2:
Next i
nstd(kk) = Sqr(nstd(kk) / (nrint - exclud2))

Case Else
'intermediar
nvalmed = 0: mvalmed = 0
exclud1 = 0 'un numarator care imi arata de cate ori am adaugat la medie
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(kk) - r) < (excl * r) Then exclud1 = exclud1 + 1: GoTo salt3
nnval(i, kk) = (r * (1 - r) * (coef(2, kk) + 2 * coef(3, kk) * r + 3 * coef(4, kk) * r ^ 2 + 4 * coef(5, kk) * r ^ 3 + 5 * coef(6, kk) * r ^ 4 + 6 * coef(7, kk) * r ^ 5 + 7 * coef(8, kk) * r ^ 6 + 8 * coef(9, kk) * r ^ 7 + 9 * coef(10, kk) * r ^ 8 + 10 * coef(11, kk) * r ^ 9) * enmed / 8.31 / ((igrec(coef(1, kk), coef(2, kk), coef(3, kk), coef(4, kk), coef(5, kk), coef(6, kk), coef(7, kk), coef(8, kk), coef(9, kk), coef(10, kk), coef(11, kk), 0.5)) ^ 2)) / (r + (r - 1) * (alphamin(kk) / (1 - alphamin(kk))))
mmval(i, kk) = nnval(i, kk) * alphamin(kk) / (1 - alphamin(kk))
nvalmed = nvalmed + nnval(i, kk)
mvalmed = mvalmed + mmval(i, kk)
salt3:
Next i
nval(kk) = nvalmed / (nrint - exclud1)
mval(kk) = mvalmed / (nrint - exclud1)
mval(kk) = nval(kk) * alphamin(kk) / (1 - alphamin(kk))
'calculez abaterile standard

mstd(kk) = 0: nstd(kk) = 0
exclud1 = 0 'un numarator care imi arata de cate ori am adaugat la medie
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(kk) - r) < (excl * r) Then exclud1 = exclud1 + 1: GoTo saltstd3
nstd(kk) = nstd(kk) + (nval(kk) - nnval(i, kk)) * (nval(kk) - nnval(i, kk))
mstd(kk) = mstd(kk) + (mval(kk) - mmval(i, kk)) * (mval(kk) - mmval(i, kk))
saltstd3:
Next i
nstd(kk) = Sqr(nstd(kk) / (nrint - exclud1))
mstd(kk) = Sqr(mstd(kk) / (nrint - exclud1))

End Select
Next kk
'calculez nre valori ale lui A, notat AP caci A parca l-am mai folosit

For kk = 1 To 2
astd(kk) = 0
amed(kk) = 0
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
ap(i, kk) = cc(kk) / (r ^ nval(kk) * (1 - r) ^ mval(kk)) / (Exp(-(enmed / 8.31 / tint(i, kk))))
amed(kk) = amed(kk) + ap(i, kk)
Next i
amed(kk) = amed(kk) / nrint '

For i = 1 To nrint
astd(kk) = astd(kk) + (amed(kk) - ap(i, kk)) * (amed(kk) - ap(i, kk))
Next i
astd(kk) = Sqr(astd(kk) / nrint)
Next kk

tiparire:
txt = vbCrLf & linie
txt = txt & vbCrLf & "Polynom degree : " & CStr(grpol - 1)
txt = txt & vbCrLf & "Exclusion domain :" & Format$(excl, "0.0####")
txt = txt & vbCrLf & "Number of intervals in the common alpha domain : " & CStr(nrint)
txt = txt & vbCrLf & linie

txt = txt & vbCrLf & "First TG curve : " & CStr(ipoints1) & " values"
txt = txt & vbCrLf & "Decomposition rate : " & Format$(cc(1), "0.0000E+00  ") & "1/sec"
txt = txt & vbCrLf & " alpha,   temp. /C,   comp. temp. /C,   deviation"
For i = 1 To ipoints1
t = dtempk(i) - 273.15
txt = txt & vbCrLf & Format$(dalpha(i), "0.0000    ") & Format$(t, "###0.00000    ") & Format$((igrec(coef(1, 1), coef(2, 1), coef(3, 1), coef(4, 1), coef(5, 1), coef(6, 1), coef(7, 1), coef(8, 1), coef(9, 1), coef(10, 1), coef(11, 1), dalpha(i)) - 273.15), "####.00000    ") & Format$(Abs(t + 273.15 - igrec(coef(1, 1), coef(2, 1), coef(3, 1), coef(4, 1), coef(5, 1), coef(6, 1), coef(7, 1), coef(8, 1), coef(9, 1), coef(10, 1), coef(11, 1), dalpha(i))), "#0.00000  ")
Next i
txt = txt & vbCrLf & linie
txt = txt & vbCrLf & "Interpolation polynom (coefficients) :"
For i = 1 To grpol
txt = txt & vbCrLf & "C" & CStr(i - 1) & "   " & Format$(coef(i, 1), "######0.0######  ")
Next i
txt = txt & vbCrLf & linie
txt = txt & vbCrLf & "Second TG curve: " & CStr(ipoints2) & " values"
txt = txt & vbCrLf & "Decomposition rate :" & Format$(cc(2), "0.0000E+00  ") & "1/sec"
txt = txt & vbCrLf & " alpha,   temp. /C,   comp. temp. /C,   deviation"

For i = ipoints1 + 2 To ipoints2 + ipoints1 + 1
t = dtempk(i) - 273.15
txt = txt & vbCrLf & Format$(dalpha(i), "0.0000    ") & Format$(t, "###0.00000    ") & Format$((igrec(coef(1, 2), coef(2, 2), coef(3, 2), coef(4, 2), coef(5, 2), coef(6, 2), coef(7, 2), coef(8, 2), coef(9, 2), coef(10, 2), coef(11, 2), dalpha(i)) - 273.15), "####.00000    ") & Format$(Abs(t + 273.15 - igrec(coef(1, 2), coef(2, 2), coef(3, 2), coef(4, 2), coef(5, 2), coef(6, 2), coef(7, 2), coef(8, 2), coef(9, 2), coef(10, 2), coef(11, 2), dalpha(i))), "#0.00000  ")
Next i
txt = txt & vbCrLf & linie
txt = txt & vbCrLf & "Interpolation polynom (coefficients) :"
For i = 1 To grpol
txt = txt & vbCrLf & "C" & CStr(i - 1) & "    " & Format$(coef(i, 1), "######0.0######")
Next i
txt = txt & vbCrLf & linie
scrie_log txt
scrie_log "Common alpha values :"
scrie_log Format$(min(3), "0.0000#     ") & Format$(max(3), "0.0000#")
scrie_log linie & vbCrLf & "Alpha for minimum temperature: curve 1,  curve 2 :"
scrie_log Format$(alphamin(1), "0.0000#     ") & Format$(alphamin(2), "0.0000#     ")
scrie_log "Minimum temperature : curve 1,  curve 2"
scrie_log Format$(tempmin(1) - 273.15, "###0.0000##     ") & Format$(tempmin(2) - 273.15, "###0.0000##")
scrie_log linie
scrie_log "Derivative : curve 1,   curve 2"
scrie_log Format$(derivata(1), "0.0000 E+00    ") & Format$(derivata(2), "0.0000 E+00    ")
scrie_log linie
scrie_log "Alpha, activation energy <J/mole>, deviation <J/mole> :"
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
scrie_log Format$(r, "0.0000     ") & Format$(en(i), "######.00    ") & Format$(Abs(en(i) - enmed), "#####.00")
Next i
scrie_log linie
scrie_log "Activation energy :" & Format$(enmed, "######.00") & "<J/mole>, std. dev. <J/mole>: " & Format$(estddev#, "#####.0")
scrie_log linie

alphaminkk = alphamin(1)
If cc(1) > cc(2) Then alphaminkk = alphamin(2)
Select Case alphaminkk '-------------------------------
Case max(3)
scrie_log "The value of n is 0."
scrie_log "Alpha,  m,  m deviation (curve 1, curve 2) :"
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(1) - r) < (excl * r) Or Abs(alphamin(2) - r) < (excl * r) Then GoTo salt4
scrie_log Format$(r, "0.0000   ") & Format$(mmval(i, 1), "0.00000  " & Format$(Abs(mmval(i, 1) - mstd(1)), "0.000   ") & Format$(mmval(i, 2), "0.00000 ") & Format$(Abs(mmval(i, 2) - mstd(2)), "0.000"))
salt4:
Next i
scrie_log "Mean value for m, m std. dev. (curve 1, curve 2) :"
scrie_log Format$(mval(1), "0.00000  ") & Format$(mstd(1), "0.000    ") & Format$(mval(2), "0.00000  ") & Format$(mstd(2), "0.000")
Case min(3)
scrie_log "The value of m is 0."
scrie_log "Alpha,  n ,  n deviation (curve 1, curve 2) :"
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(1) - r) < (excl * r) Or Abs(alphamin(2) - r) < (excl * r) Then GoTo salt5
scrie_log Format$(r, "0.0000   ") & Format$(nnval(i, 1), "0.00000  ") & Format$(Abs(nnval(i, 1) - nstd(1)), "0.000    ") & Format$(nnval(i, 2), "0.00000  ") & Format$(Abs(nnval(i, 2) - nstd(2)), "0.000")
salt5:
Next i
scrie_log "Mean value for n, n std. dev. (curve 1, curve 2) :"
scrie_log Format$(nval(1), "0.00000  ") & Format$(nstd(1), "0.000    ") & Format$(nval(2), "0.00000  ") & Format$(nstd(2), "0.000   ")

Case Else
'intermediar

scrie_log "Alpha,  n,  deviation (curve 1, curve 2) :"
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
If Abs(alphamin(1) - r) < (excl * r) Or Abs(alphamin(2) - r) < (excl * r) Then GoTo salt6
scrie_log Format$(r, "0.0000   ") & Format$(nnval(i, 1), "0.00000  ") & Format$(Abs(nnval(i, 1) - nval(1)), "0.000    ") & Format$(nnval(i, 2), "0.00000  ") & Format$(Abs(nnval(i, 2) - nval(2)), "0.000")
salt6:
Next i
scrie_log linie
scrie_log "Mean  n,  n. std. dev. (curve 1, curve 2): "
scrie_log Format$(nval(1), "0.00000  ") & Format$(nstd(1), "0.000    ") & Format$(nval(2), "0.00000  ") & Format$(nstd(2), "0.000")
scrie_log linie
scrie_log "Mean  m, m std. dev. (curve 1, curve 2): "
scrie_log Format$(mval(1), "0.00000  ") & Format$(mstd(1), "0.000    ") & Format$(mval(2), "0.00000  ") & Format$(mstd(2), "0.000")
End Select
'''scrie_log ( linie$
scrie_log linie
scrie_log "Alpha, preexponential factor, deviation (curve 1, curve 2) :"
For i = 1 To nrint
r = min(3) + (i - 1) * (max(3) - min(3)) / nrint
scrie_log Format$(r, "0.0000  ") & Format$(ap(i, 1), "0.00000 E+00  ") & Format$(Abs(ap(i, 1) - amed(1)), "0.000 E+00    ") & Format$(ap(i, 2), "0.00000 E+00   ") & Format$(Abs(ap(i, 2) - amed(2)), "0.000 E+00")
Next i
scrie_log linie
scrie_log "Preexponential factor  and std. dev. (curve 1, curve 2), <1/sec> :"
scrie_log Format$(amed(1), "0.00000 E+00  ") & Format$(astd(1), "0.000 E+00       ") & Format$(amed(2), "0.00000 E+00  ") & Format$(astd(2), "0.000 E+00")
scrie_log linie
gindicator(3) = True
Exit Sub
handleit:
t = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
Close
Exit Sub
End Sub




Sub curat_grid()
On Error GoTo handleit
For j% = 0 To 2
data_editor.Grid1.Col = j%
For i% = 1 To data_editor.Grid1.Rows - 1
data_editor.Grid1.Row = i%
data_editor.Grid1.Text = ""
Next i%
Next j%
data_editor.Grid1.Rows = 21
Exit Sub
handleit:
MsgBox "Unexpected error in curat_grid routine." & vbCrLf & CStr(Err.Description)
Exit Sub
End Sub

Sub ordon(inceput As Integer, sfarsit As Integer, coloana As Integer, eroare As Boolean)
'ordoneaza gridul intre liniile inceput si sfarsit, inclusiv; dupa coloana
'eroare as boolean
On Error GoTo handleit
    data_editor!Grid1.Col = coloana
   Do
    itest% = 0
    For i% = inceput To sfarsit - 1
    data_editor!Grid1.Row = i%
    sa# = Val(data_editor!Grid1.Text)
    data_editor!Grid1.Row = i% + 1
    sb# = Val(data_editor!Grid1.Text)
    If sb# < sa# Then
    data_editor!Grid1.Text = sa#
    data_editor!Grid1.Row = i%
    data_editor!Grid1.Text = sb#
    itest% = 1
    End If
    Next i%
   Loop Until itest% = 0
Exit Sub
handleit:
eroare = True
Exit Sub
End Sub



Sub deriv(n As Integer, y() As Double, ddy() As Double, eroare As Boolean)
'calculeaza derivata numerica prin derivata polinomului de interpolare
'este posibil ca eroarea sa fie mare
On Error GoTo localhandle
ddy(1) = 0
ddy(n) = 0
'Dim coef() As Double
'ReDim coef(n, 3)
'Dim aa(5, 3) As Double, ii(5) As Double, xx(3) As Double
'For j = 3 To n - 2
'aa(1, 1) = 1: aa(1, 2) = dx(j - 2): aa(1, 3) = (dx(j - 2)) ^ 2 ': aa(1, 4) = (dx(j - 2)) ^ 3
'aa(2, 1) = 1: aa(2, 2) = dx(j - 1): aa(2, 3) = (dx(j - 1)) ^ 2 ': aa(2, 4) = (dx(j - 1)) ^ 3
'aa(3, 1) = 1: aa(3, 2) = dx(j): aa(3, 3) = (dx(j)) ^ 2 ': aa(3, 4) = (dx(j)) ^ 3
'aa(4, 1) = 1: aa(4, 2) = dx(j + 1): aa(4, 3) = (dx(j + 1)) ^ 2 ': aa(4, 4) = (dx(j + 1)) ^ 3
'aa(5, 1) = 1: aa(5, 2) = dx(j + 2): aa(5, 3) = (dx(j + 2)) ^ 2 ': aa(5, 4) = (dx(j + 2)) ^ 3
'ii(1) = y(j - 2): ii(2) = y(j - 1): ii(3) = y(j): ii(4) = y(j + 1): ii(5) = y(j + 2)
'Call pseudoinv(5, 3, aa(), ii(), xx(), lowsize, eroare)
'If eroare Then Err.Raise 1101
'ddy(j) = xx(2) + 2 * xx(3) * dx(j)
'coef(j, 1) = xx(1): coef(j, 2) = xx(2): coef(j, 3) = xx(3): coef(j, 4) = xx(4)
'Next j
'i = 2
For i = 2 To n - 1
ddy(i) = y(i - 1) * (dx(i) - dx(i + 1)) / (dx(i - 1) - dx(i)) / (dx(i - 1) - dx(i + 1)) + y(i) * (2 * dx(i) - dx(i - 1) - dx(i + 1)) / (dx(i) - dx(i - 1)) / (dx(i) - dx(i + 1)) + y(i + 1) * (dx(i) - dx(i - 1)) / (dx(i + 1) - dx(i - 1)) / (dx(i + 1) - dx(i))
Next i
'i = n - 1
'ddy(i) = y(i - 1) * (dx(i) - dx(i + 1)) / (dx(i - 1) - dx(i)) / (dx(i - 1) - dx(i + 1)) + y(i) * (2 * dx(i) - dx(i - 1) - dx(i + 1)) / (dx(i) - dx(i - 1)) / (dx(i) - dx(i + 1)) + y(i + 1) * (dx(i) - dx(i - 1)) / (dx(i + 1) - dx(i - 1)) / (dx(i + 1) - dx(i))
Exit Sub
localhandle:
eroare = True
Exit Sub
End Sub

Sub scrie_log(mesaj As String)
main_display.richtxtlog.Text = main_display.richtxtlog.Text & mesaj + vbCrLf
DoEvents
End Sub
Sub verif_grid(inceput As Integer, sfarsit As Integer, puncte As Integer)
On Error GoTo handle
    itest% = inceput
   Do
    data_editor!Grid1.Col = 1
    itest% = itest% + 1
    data_editor!Grid1.Row = itest%
    If data_editor!Grid1.Row = sfarsit Then puncte = itest% - inceput: Exit Do
    sa# = Val(data_editor!Grid1.Text)
    data_editor!Grid1.Col = 2
    sb# = Val(data_editor!Grid1.Text)
    If (sa# <= 0) And (sb# <= 0) Then puncte = itest% - inceput: Exit Do
   Loop
    puncte = itest% - inceput
Exit Sub
handle:
MsgBox "There may be an error (module.verif_grid routine)..."
Exit Sub
End Sub

Sub numerotez(inceput As Integer, sfarsit As Integer)
On Error GoTo handle
    data_editor.Grid1.Col = 0
    For i% = inceput To sfarsit
    data_editor.Grid1.Row = i%
    data_editor.Grid1.Text = CStr(i%)
    Next i%
    Exit Sub
handle:
Exit Sub
End Sub

Sub citesc(inceput As Integer, sfarsit As Integer, er As Boolean)
On Error GoTo handleit
er = False
    For i% = inceput To sfarsit
    data_editor!Grid1.Col = 1
    data_editor!Grid1.Row = i%
    dx(i%) = Val(data_editor!Grid1.Text)
    data_editor!Grid1.Col = 2
    dy(i%) = Val(data_editor!Grid1.Text)
    If dy(i%) < 0 Then Err.Raise 1101
    If (dx(i%) <= 0) And (dy(i%) <= 0) Then Err.Raise 1101
    Next i%
Exit Sub
handleit:
eroare = True
MsgBox ("Error reading grid.")
Exit Sub
End Sub

Sub verif_temp(inceput As Integer, sfarsit As Integer, eroare As Boolean)
On Error GoTo handleit
        Select Case data_editor.lst1.ListIndex
        Case 0 'grade C
        For i% = inceput To sfarsit
        dtempk(i%) = dx(i%) + 273.15
        Next i%
        Case 1 '  grade K
        For i% = inceput To sfarsit
        dtempk(i%) = dx(i%)
        Next i%
        End Select
        
        For i% = inceput To sfarsit
        If (dtempk(i%) < 233 Or dtempk(i%) > 1700) Then Err.Raise 1101
        Next i%
eroare = False
Exit Sub
handleit:
eroare = True
Exit Sub
End Sub
Sub verif_y(inceput As Integer, sfarsit As Integer, eroare As Boolean)
On Error GoTo handleit
'verifica valorile pentru dy(i), deja citite
    Select Case data_editor.lst2.ListIndex
    Case 0 'alpha
        If ((Not (dy(1) = 0)) Or (Not (dy(ipoints) = 1))) Then Err.Raise 1101 ' 1101, , "The alpha values for the first and the last point must be 0 and 1"
        For i% = inceput + 1 To sfarsit - 1
        If (dy(i%) < 0 Or dy(i%) > 1) Then Err.Raise 1101 ', , "Incorrect alpha values, accepted: between 0 and 1"
        Next i%
    Case 1, 2 'dta, dtg
    Case 3 'tg
        For i% = inceput + 1 To sfarsit - 1
        If dy(i%) > dy(ipoints) Then Err.Raise 1101 ', , "There are errors in TG values. The last point expected for alpha=1."
        Next i%
    End Select
eroare = False
Exit Sub
handleit:
eroare = True
Exit Sub
End Sub

Sub verif_data(eroare As Boolean)
'subrutina de verificare a datelor, intoarce un mesaj de eroare in caz de probleme
'intoarce eroare in caz de probleme, dar si cu msgbox aici
'daca iese bine de aici eroare = False
On Error GoTo handleit
eroare = False
nume_exp = data_editor.txtname.Text
If Len(data_editor.txtstd.Text) < 1 Then Err.Raise 1101, , "Input the deviation , %."
sigy = CDbl(data_editor.txtstd.Text)
If (sigy < 0.0005 Or sigy > 20) Then Err.Raise 1101, , "The estimated error of your data is " & Format$(data_editor.txtstd.Text, "0.###") & "?" & vbCrLf & "(accepted values:between 0.0005% and 20%)"
Select Case data_editor.tabdata.Caption
Case "Integral"
    Call verif_grid(1, data_editor.Grid1.Rows - 1, ipoints)
    If ipoints < 7 Then Err.Raise 1101, , "I found " & CStr(ipoints) & " pairs of data. I need at least 7."
ReDim Preserve dx(ipoints)
ReDim Preserve dy(ipoints)
ReDim Preserve dtempk(ipoints)
ReDim Preserve dalpha(ipoints)
  
    Call numerotez(1, ipoints) ''numerotez coloanele
    Call ordon(1, ipoints, 1, eroare)    'ordonez valorile
    If eroare Then Err.Raise 1101, , "Error processing the data points in grid. Check your data."
    Call citesc(1, ipoints, eroare) 'citesc valorile din grid in dx, dy si valorile temperaturilor
    If eroare Then Err.Raise 1101, , "Error trying to read the grid (I accept only positive values)."
    Call verif_temp(1, ipoints, eroare)
    If eroare Then Err.Raise 1101, , "Incorrect values of temperature. Check your data."
    Call verif_y(1, ipoints, eroare)
    If eroare Then Err.Raise 1101, , "Incorrect values for alpha or TG. Check your data."
    itest% = data_editor.chkint(0).Value + data_editor.chkint(1).Value + data_editor.chkint(2).Value + data_editor.chkint(3).Value
    If itest% = 0 Then Err.Raise 1101, , "You forget to select the methods."
    'verifica ninitial, nfinal, etc si le atribui aici
    ninitial = Val(data_editor.txtint(0).Text)
    nfinal = Val(data_editor.txtint(1).Text)
    nstep = Val(data_editor.txtint(2).Text)
    rate = Val(data_editor.txtint(3).Text)
    If ninitial < 0 Or ninitial > 3 Then Err.Raise 1101, , "Initial reaction order is " & CStr(ninitial) & " ?" & vbCrLf & "Accepted values: between 0 and 3."
    If nfinal < 0 Or nfinal > 3 Then Err.Raise 1101, , "Final reaction order is " & CStr(nfinal) & " ?" & vbCrLf & "Accepted values: lower than 3."
    If nfinal < ninitial Then Err.Raise 1101, , "Final reaction order is bigger than the initial one."
    If nstep < 0.001 Or nstep > (nfinal - ninitial) / 3 Then Err.Raise 1101, , "Incorrect reaction order step."
    If rate < 0.001 Or rate > 50 Then Err.Raise 1101, , "Heating rate is " & CStr(rate) & " K/min ???. Accepted values: between 0.05 and 50 K/min."
    If (nfinal - ninitial) / nstep > 250 Then MsgBox "You may have computing problems. Too many steps, i'll try..."
Case "Differential"
    Call verif_grid(1, data_editor.Grid1.Rows - 1, ipoints)
    If ipoints < 7 Then Err.Raise 1101, , "I found " & CStr(ipoints) & " pairs of data. I need at least 7."
ReDim dx(ipoints), dy(ipoints), dtempk(ipoints), dalpha(ipoints)
   
    Call numerotez(1, ipoints) ''numerotez coloanele
    Call ordon(1, ipoints, 1, eroare)    'ordonez valorile
    If eroare Then Err.Raise 1101, , "Error processing the data points in grid. Check your data."
    Call citesc(1, ipoints, eroare) 'citesc valorile din grid in dx, dy si valorile temperaturilor
    If eroare Then Err.Raise 1101, , "Error trying to read the grid (I accept only positive values)."
    Call verif_temp(1, ipoints, eroare)
    If eroare Then Err.Raise 1101, , "Incorrect values of temperature. Check your data."
    Call verif_y(1, ipoints, eroare)
    If eroare Then Err.Raise 1101, , "Incorrect values for alpha or TG. Check your data."
    itest% = data_editor.chkdif(0).Value + data_editor.chkdif(1).Value + data_editor.chkdif(2).Value + data_editor.chkdif(3).Value
    If itest% = 0 Then Err.Raise 1101, , "You forget to select the methods."
    'verifica ninitial, nfinal, etc si le atribui aici
    ninitial = Val(data_editor.txtdif(0).Text)
    nfinal = Val(data_editor.txtdif(1).Text)
    nstep = Val(data_editor.txtdif(2).Text)
    rate = Val(data_editor.txtdif(3).Text)
    If ninitial < 0 Or ninitial > 3 Then Err.Raise 1101, , "Initial reaction order is " & CStr(ninitial) & " ?" & vbCrLf & "Accepted values: between 0 and 3."
    If nfinal < 0 Or nfinal > 3 Then Err.Raise 1101, , "Final reaction order is " & CStr(nfinal) & " ?" & vbCrLf & "Accepted values: lower than 3."
    If nfinal < ninitial Then Err.Raise 1101, , "Final reaction order is bigger than the initial one."
    If nstep < 0.001 Or nstep > (nfinal - ninitial) / 3 Then Err.Raise 1101, , "Incorrect reaction order step."
    If rate < 0.001 Or rate > 50 Then Err.Raise 1101, , "Heating rate is " & CStr(rate) & " K/min ???. Accepted values: between 0.05 and 50 K/min."
    If (nfinal - ninitial) / nstep > 1000 Then MsgBox "You may have computing problems. Too many steps, i'll try..."

Case "Regression"
    Call verif_grid(1, data_editor.Grid1.Rows - 1, ipoints)
    If ipoints < 7 Then Err.Raise 1101, , "I found " & CStr(ipoints) & " pairs of data. I need at least 7."
    ReDim dx(ipoints), dy(ipoints), dtempk(ipoints), dalpha(ipoints)
    Call numerotez(1, ipoints) ''numerotez coloanele
    Call ordon(1, ipoints, 1, eroare)    'ordonez valorile
    If eroare Then Err.Raise 1101, , "Error processing the data points in grid. Check your data."
    Call citesc(1, ipoints, eroare) 'citesc valorile din grid in dx, dy si valorile temperaturilor
    If eroare Then Err.Raise 1101, , "Error trying to read the grid (I accept only positive values)."
    Call verif_temp(1, ipoints, eroare)
    If eroare Then Err.Raise 1101, , "Incorrect values of temperature. Check your data."
    Call verif_y(1, ipoints, eroare)
    If eroare Then Err.Raise 1101, , "Incorrect values for alpha or TG. Check your data."
    rate = CDbl(data_editor!txtreg.Text)
    If rate < 0.001 Or rate > 50 Then Err.Raise 1101, , "Heating rate is " & CStr(rate) & " K/min ???. Accepted values: between 0.05 and 50 K/min."
'    itest% = data_editor!chkreg(3).Value + data_editor!chkreg(4).Value + data_editor!chkreg(5).Value
'    If itest% = 0 Then Err.Raise 1101, , "You forget to select the regression method."
    itest% = data_editor!chkreg(0).Value + data_editor!chkreg(1).Value + data_editor!chkreg(2).Value
    If itest% = 0 Then Err.Raise 1101, , "You forget to select the conversion function."
    If itest% > 2 Then Err.Raise 1101, , "Conversion function not accepted. Select any combination but no more than two parameters."

Case "CRTA"
'verific datele din crta, ipoints1 si ipoints2 ca globale pentru crta 2, ca si rate1, rate2
'daca nu este crta2 gridul este clasic, ordonarea se face dupa alpha
    'crta 2
    Call verif_grid(1, data_editor.Grid1.Rows - 1, ipoints1)
    If ipoints1 < 6 Then Err.Raise 1101, , "For the 1st data set I found " & CStr(ipoints1) & " pairs of data. You need at least 6."
    Call verif_grid(ipoints1 + 2, data_editor.Grid1.Rows - 1, ipoints2)
    If ipoints2 < 6 Then Err.Raise 1101, , "For the 2nd data set I found " & CStr(ipoints2) & " pairs of data. You need at least 6."
    ReDim dx(ipoints1 + ipoints2 + 1), dy(ipoints1 + ipoints2 + 1), dtempk(ipoints1 + ipoints2 + 1), dalpha(ipoints1 + ipoints2 + 1)
    Call numerotez(1, ipoints1) ''numerotez coloanele
    Call numerotez(ipoints1 + 2, ipoints1 + ipoints2 + 1)
    Call ordon(1, ipoints1, 2, eroare)
    If eroare Then Err.Raise 1101, , "Error processing the 1st data set in grid. Check your data."
    Call ordon(ipoints1 + 2, ipoints2 + ipoints1 + 1, 2, eroare)
    If eroare Then Err.Raise 1101, , "Error processing the 2nd data set in grid. Check your data."
'-------
'citesc valorile din grid in dx, dy si valorile temperaturilor
    Call citesc(1, ipoints1, eroare) 'citesc valorile din grid in dx, dy si valorile temperaturilor
    If eroare Then Err.Raise 1101, , "Error trying to read the grid (I accept only positive values)."
    Call citesc(ipoints1 + 2, ipoints1 + ipoints2 + 1, eroare) 'citesc valorile din grid in dx, dy si valorile temperaturilor
    If eroare Then Err.Raise 1101, , "Error trying to read the grid (I accept only positive values)."
    Call verif_temp(1, ipoints1, eroare)
    Call verif_temp(ipoints1 + 2, ipoints1 + ipoints2 + 1, eroare)
    If eroare Then Err.Raise 1101, , "Incorrect value of temperature."
    'verifica valorile pentru dy(i), deja citite
    ivm1 = Val(data_editor.txtcrta(1).Text)
    ivm2 = Val(data_editor.txtcrta(2).Text)
    iivm1 = Val(data_editor.txtcrta(4).Text)
    iivm2 = Val(data_editor.txtcrta(5).Text)
        Select Case data_editor.lst2.ListIndex
        Case 0 'alpha
            For i% = 1 To ipoints1
            If (dy(i%) < 0 Or dy(i%) > 1) Then Err.Raise 1101, , "Incorrect alpha values, accepted: between 0 and 1"
            dalpha(i%) = dy(i%)
            Next i%
            For i% = ipoints1 + 2 To ipoints1 + ipoints2
            If (dy(i%) < 0 Or dy(i%) > 1) Then Err.Raise 1101, , "Incorrect alpha values, accepted: between 0 and 1"
            dalpha(i%) = dy(i%)
            Next i%
        Case 1, 2 'dta, dtg
        Err.Raise 1101, , "I take alpha or TG values only."
        Case 3 'tg
        For i% = 1 To ipoints1
            If dy(i%) > ivm2 Then Err.Raise 1101, , "There are errors in TG values. Check the TG for alpha=1 for the first data set."
            If dy(i%) < ivm1 Then Err.Raise 1101, , "There are errors in TG values. Check the TG for alpha=0 for the first data set."
            dalpha(i%) = (dy(i%) - ivm1) / (ivm2 - ivm1)
        Next i%
        For i% = ipoints1 + 2 To ipoints2 + 1
            If dy(i%) > iivm2 Then Err.Raise 1101, , "There are errors in TG values. Check the TG for alpha=1 for the 2nd data set."
            If dy(i%) < iivm1 Then Err.Raise 1101, , "There are errors in TG values. Check the TG for alpha=0 for the 2nd data set."
            dalpha(i%) = (dy(i%) - iivm1) / (iivm2 - iivm1)
        
        Next i%
    End Select
For i% = 0 To 8
If Len(data_editor!txtcrta(i%).Text) = 0 Then Err.Raise 1101, , "A text field is empty : " & data_editor.lblcrta(i%).Caption & vbCrLf & "Input a valid value and then try again."
Next i%
rate1 = CDbl(data_editor!txtcrta(0).Text)
rate2 = CDbl(data_editor!txtcrta(3).Text)
If rate1 < 0.0000005 Or rate1 > 10 Then Err.Raise 1101, , "Rate is " & CStr(rate) & " mg/min ???. Accepted values: between 0.0005 and 10 mg/min."
If rate2 < 0.0000005 Or rate2 > 10 Then Err.Raise 1101, , "Rate is " & CStr(rate1) & " mg/min ???. Accepted values: between 0.0005 and 10 mg/min."
If rate1 = rate2 Then Err.Raise 1101, , "You need different decomposition rates for ICAR 2."
If data_editor.lst2.ListIndex = 0 Then
data_editor.txtcrta(1).Text = "0.0": data_editor.txtcrta(2).Text = "1.0"
Else
If (data_editor.txtcrta(1).Text > data_editor.txtcrta(2).Text) Then Err.Raise 1101, , "Check the values for alpha=0 and 1."
If data_editor.txtcrta(1).Text < 0 Then Err.Raise 1101, , "I want positive values only."
End If

If data_editor.lst2.ListIndex = 0 Then
data_editor.txtcrta(4).Text = "0.0": data_editor.txtcrta(5).Text = "1.0"
Else
If (data_editor.txtcrta(4).Text > data_editor.txtcrta(5).Text) Then Err.Raise 1101, , "Check the values for alpha=0 and 1."
If data_editor.txtcrta(4).Text < 0 Then Err.Raise 1101, , "I want positive values only."
End If
nrint = CInt(data_editor.txtcrta(7).Text)
excl = CDbl(data_editor.txtcrta(6).Text)
grpol = CInt(data_editor.txtcrta(8).Text) + 1
If excl < 0.0001 Or excl > 0.15 Then Err.Raise 1101, , "Inconsistent excluded domain, see the help file."
If nrint < 10 Or nrint > 1000 Then Err.Raise 1101, , "Wrong number of intervals.(it must be between 10 and 1000)."
If grpol < 2 Or grpol > 11 Then Err.Raise 1101, , "The accepted polynom degree is between 2 and 8, inclusive."
If ipoints1 < grpol Or ipoints2 < grpol Then Err.Raise 1101, , "Not enough data points for interpolation !"


End Select
eroare = False 'daca a ajuns aici
Exit Sub

handleit:
i% = MsgBox(CStr(Err.Description), vbOKOnly + vbInformation, nume_prog)
eroare = True
Exit Sub 'exit sub face el err.clear
End Sub

Sub surfint(n As Integer, x() As Double, y() As Double, al() As Double, eroare As Boolean)
Dim i As Integer, s As Double
eroare = False
On Error GoTo localhandle
al(1) = 0
s = 0
For j = 2 To n
s = 0
For i = 2 To j
s = s + ((y(i) + y(i - 1)) / 2) * (x(i) - x(i - 1))
Next i
al(j) = s
Next j
For i = 2 To n - 1
al(i) = al(i) / al(n)
Next i
al(n) = 1
eroare = False
Exit Sub
localhandle:
eroare = True
Exit Sub
End Sub

Sub exp_points()
On Error GoTo handle
'rutina asta numara cate perechi de puncte exista, ipoints, in grid
Dim itest As Integer, sa As Single, sb As Single
'nu citesc valorile
itest = 0
Do
data_editor!Grid1.Col = 1
itest = itest + 1
data_editor!Grid1.Row = itest
sa = Val(data_editor!Grid1.Text)
data_editor!Grid1.Col = 2
sb = Val(data_editor!Grid1.Text)
If (sa <= 0) And (sb <= 0) Then Exit Do
Loop
ipoints = itest - 1
Exit Sub

handle:
Exit Sub
End Sub


Sub reglin(ns As Integer, x() As Double, y() As Double, sig As Double, ms As Double, bs As Double, cor As Double, dm As Double, db As Double, eroares As Boolean)
On Error GoTo handleit
'regresie ponderata cu x si y, acesta are dev std. sig
'atentie: sig este eroare procentuala la calcul, trebuie calculata pentru fiecare liny...
'Subrutina intoarce coeficientul'de corelatie, panta, ordonata cu erorile lor
's vine de la apel de subrutina, de tip privat
'eroares este un indice de eroare, daca este 1 atunci arata insucces
'ns este numarul de puncte
'm este panta iar b este ordonata
eroares = False
'Dim stest As Double
sigma = 2 * 0.141 * sig * y(CInt((ipoints - 2) / 2))
sx = 0: sy = 0: sxy = 0: sxx = 0: syy = 0: stest = 0
For i = 2 To ns - 1
sx = sx + x(i): sy = sy + y(i): sxx = sxx + x(i) * x(i): sxy = sxy + x(i) * y(i): syy = syy + y(i) * y(i)
Next i
ms = (sxy - sx * sy / (ns - 2)) / (sxx - sx * sx / (ns - 2))
bs = sy / (ns - 2) - ms * sx / (ns - 2)
'MsgBox ((ns - 2) * (sxx - sx * sx))
dm = Sqr(Abs((sigma * sigma) / ((ns - 2) * (sxx - sx * sx))))
'MsgBox CStr(dm)
db = Sqr(Abs((sigma * sigma * sxx) / ((ns - 2) * (sxx - sx * sx))))
cor = (sxy - sx * sy / (ns - 2)) / Sqr(Abs((sxx - sx * sx / (ns - 2)) * (syy - sy * sy / (ns - 2))))
'sqchi, calcul
'sqchi = 0
'For i = 2 To ns - 1
'sqchi = sqchi + (Y(i) - ms * X(i) - bs) * (Y(i) - ms * X(i) - bs) / (sig * sig)
'Next i
eroares = False
Exit Sub
handleit:
eroares = True
Exit Sub
End Sub

Sub deschide_fisier(ByVal nume_fisier As String, ByVal intrare_iesire_edit As Integer, filtru As String, indice As Integer)
'nume_fisier trebuie sa fie cu path, respectand regulile de scriere DOS
'scriere si citire din directorul curent
'daca intrare_iesire=1 inseamna input, trebuie ca fisierul sa existe
'daca intrare_iesire=2 inseamna output, trebuie ca fisierul sa fie nou,
'eventual intreaba de supra scriere
'trateaza eroarea eventuala la scriere si citire
'--------------------------------------------------
Select Case intrare_iesire_edit
    Case 1 'citirea path_input_text
On Error GoTo error_open_1
main_display.comdialog1.Filter = filtru
main_display.comdialog1.FilterIndex = indice
main_display.comdialog1.Flags = &H1000& Or &H4& Or &H800&
'ofn_filemustexist
'ofn_readonly
'ofn_pathmustexist
main_display.comdialog1.DialogTitle = nume_prog & " - Input file"
main_display.comdialog1.ShowOpen
'main_display.inputfile = comdialog1.filename
'form1.Panel3D1.ForeColor = 0
'form1.Panel3D1.Caption = Right$(LCase$(comdialog1.filename), 30)
no_input = False
inputfile = main_display.comdialog1.filename

    Case 2 'citire path_output_text
On Error GoTo error_open_2
main_display.comdialog1.Filter = filtru
main_display.comdialog1.FilterIndex = indice
main_display.comdialog1.filename = ""
main_display.comdialog1.Flags = &H2& Or &H1& Or &H800& Or &H4&
'ofn_overwriteprompt
'ofn_readonly
'ofn_pathmustexist
main_display.comdialog1.DialogTitle = nume_prog & " - Output file"
main_display.comdialog1.ShowSave
outputfile = main_display.comdialog1.filename
'form1.Panel3D2.ForeColor = &H0
'form1.Panel3D2.Caption = Right$(LCase$(comdialog1.filename), 30)
'Open comdialog1.filename For Output As #2
no_output = False
Case Else 'citire path_edit_text
       'paseaza actiunea la editor
       'imposibil
MsgBox (" Error in general_select_file ")
End Select
Screen.MousePointer = 0
Close
Exit Sub

error_open_1:
If Err.Number = 32755 Then no_input = True: Exit Sub
MsgBox ("There were errors trying to open input file - " & CStr(Err))
Screen.MousePointer = 0
Close
Exit Sub


error_open_2:
If Err = 32755 Then no_output = True: Exit Sub
MsgBox ("There were errors trying to open a file - " & CStr(Err))
Screen.MousePointer = 0
Close
Exit Sub

End Sub

Sub pseudoinv(ne As Integer, n As Integer, z() As Double, ii() As Double, x() As Double, lowsize As Double, eroare As Boolean)
'ne este numarul de ecuatii
'n este numarul de necunoscute, n<=ne
'lowsize este valoarea minima a pivotului, de ordin a 10^-10
'z este matricea coeficientilor sistemului de determinat
'i este termenul liber
'x este solutia
eroare = False
On Error GoTo handleit
Dim i As Integer, j As Integer, k As Integer, semn As Integer, lc As Integer
Dim maxx As Double, d As Double, aaA As Double, m As Integer
ReDim a(n, n) As Double
    For i = 1 To n: For j = 1 To n: a(i, j) = 0: For k = 1 To ne
        a(i, j) = a(i, j) + z(k, i) * z(k, j)
    Next k: Next j: Next i
 ReDim c(n, n) As Double
    For i = 1 To n
         c(i, i) = 1
     Next i
 semn = 1
     For m = 1 To n - 1: maxx = Abs(a(m, m)): lc = m
         For i = m + 1 To n
         If Abs(a(i, m)) > maxx Then maxx = Abs(a(i, m)): lc = i
         Next i
         If maxx < lowsize Then Err.Raise 1101
         For i = 1 To n
             d = a(lc, i): a(lc, i) = a(m, i): a(m, i) = d
             d = c(lc, i): c(lc, i) = c(m, i): c(m, i) = d
         Next i
 If lc <> m Then semn = -semn
 For i = m + 1 To n: aaA = a(i, m)
     For j = 1 To n
         a(i, j) = a(i, j) - a(m, j) * aaA / a(m, m)
         c(i, j) = c(i, j) - c(m, j) * aaA / a(m, m)
     Next j: Next i: Next m
     For m = n To 2 Step -1
         For i = m - 1 To 1 Step -1: aaA = a(i, m)
             For j = n To 1 Step -1
             a(i, j) = a(i, j) - a(m, j) * aaA / a(m, m)
             c(i, j) = c(i, j) - c(m, j) * aaA / a(m, m)
             Next j
         Next i
     Next m
 For i = 1 To n: For j = 1 To n: c(i, j) = c(i, j) / a(i, i): Next j: Next i
 ReDim bb(n, ne) As Double
 For i = 1 To n: For j = 1 To ne: For k = 1 To n: bb(i, j) = bb(i, j) + c(i, k) * z(j, k): Next k: Next j: Next i
 For i = 1 To n: x(i) = 0: For k = 1 To ne: x(i) = x(i) + bb(i, k) * ii(k): Next k: Next i
 Erase a, c, bb 's ar putea sa fie inutile, nu sunt globale
Exit Sub
handleit:
eroare = True
Exit Sub
End Sub


Sub simulate(al() As Double, te() As Double, tempstart As Double, tempend As Double, n As Double, a As Double, E As Double, r As Double, npx As Integer, eroare As Boolean)
'alpha si temp sunt locale, nu pot depasi 200 de valori
'calcule alpha in functie de temp pentru n, m si p
On Error GoTo handleit
Dim p As Double, x As Double, creste As Boolean, max As Double
max = 0
If n <= 1.001 And n > 0.999 Then n = 1.0001
dt = CDbl((tempend - tempstart) / 200)
k% = 1
t# = tempstart
Do While k% < 201
x = -E / 8.314 / t#
p = 0
For j% = 1 To npx
p = p + factorial(j%) / (x ^ (j% - 1))
Next j%
p = p * Exp(x) / x / x
alpha# = 1 - ((n - 1) * a / r * E / 8.314 * p + 1) ^ (1 / (1 - n))
te(k%) = t#
If alpha# < 0 Then alpha# = 0
If alpha# > 1 Then alpha# = 1
If alpha# > 0.1 Then creste = True
If creste And (k% > 1) And (alpha# <= max) Then alpha# = 1
al(k%) = alpha#
max = alpha#
t# = t# + dt
k% = k% + 1
Loop
eroare = False
Exit Sub
handleit:
Resume Next
'eroare = True
Exit Sub
End Sub

