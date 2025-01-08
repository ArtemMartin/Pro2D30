VERSION 5.00
Begin VB.Form OZzelkagdform 
   Caption         =   "Огневая задача Цель каждому"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   8000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3600
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Цель"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   10095
      Begin VB.ComboBox pvvNc3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5800
         TabIndex        =   29
         Text            =   "0"
         Top             =   4300
         Width           =   1500
      End
      Begin VB.ComboBox pvvNc2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3500
         TabIndex        =   28
         Text            =   "0"
         Top             =   4300
         Width           =   1500
      End
      Begin VB.ComboBox pvvNc1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1200
         TabIndex        =   27
         Text            =   "0"
         Top             =   4300
         Width           =   1500
      End
      Begin VB.TextBox pGlc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   5800
         TabIndex        =   24
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.TextBox pFrc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   5800
         TabIndex        =   23
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox phc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   5800
         TabIndex        =   22
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pYc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   5800
         TabIndex        =   21
         Text            =   "0"
         Top             =   1600
         Width           =   1500
      End
      Begin VB.TextBox pXc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   5800
         TabIndex        =   20
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pGlc2 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   3500
         TabIndex        =   19
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.TextBox pFrc2 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   3500
         TabIndex        =   18
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox phc2 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   3500
         TabIndex        =   17
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pYc2 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   3500
         TabIndex        =   16
         Text            =   "0"
         Top             =   1600
         Width           =   1500
      End
      Begin VB.TextBox pXc2 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   3500
         TabIndex        =   15
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pGlc1 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   14
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.TextBox pFrc1 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   13
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox phc1 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   12
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pYc1 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   11
         Text            =   "0"
         Top             =   1600
         Width           =   1500
      End
      Begin VB.TextBox pXc1 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   10
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "Решить"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   8000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1000
         Width           =   1700
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл Цели="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   100
         TabIndex        =   26
         Top             =   4200
         Width           =   1000
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Гл="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   9
         Top             =   3400
         Width           =   700
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фр="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   8
         Top             =   2800
         Width           =   700
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "hц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   7
         Top             =   2200
         Width           =   700
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Уц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   6
         Top             =   1600
         Width           =   700
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Хц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   5
         Top             =   1000
         Width           =   700
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   6200
         TabIndex        =   4
         Top             =   400
         Width           =   900
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   400
         Left            =   3900
         TabIndex        =   3
         Top             =   400
         Width           =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1600
         TabIndex        =   2
         Top             =   400
         Width           =   900
      End
   End
End
Attribute VB_Name = "OZzelkagdform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'записываем номера целей
Open App.Path & "\writeZeliEach" For Output As #1
Write #1, pvvNc1, pvvNc2, pvvNc3
Close #1

 Dim dDov1 As Single, Dret1 As Single, dDr1 As Single, dN As Single
 Dim rep1 As String, rep2 As String, rep3 As String
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
50:
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
   Xc1 = pXc1: Yc1 = pYc1: hc1 = phc1
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = Xc1 - Xop1
60: dy1 = Yc1 - Yop1
61: dh1 = hc1 - hop1
   Pi = 3.14159265358
9010: Dt1 = Int(Sqr(dx1 ^ 2 + dy1 ^ 2) + 0.001)
9110: Yr1 = CInt((dh1 / (Dt1 * 0.001 + 0.001)) * 0.95)
100: A1 = Abs(Atn(dy1 / (dx1 + 0.001)) / Pi * 30) * 100
101: If dx1 > 0 And dy1 > 0 Then Ygolt1 = CInt(A1)
102: If dx1 < 0 And dy1 > 0 Then Ygolt1 = CInt(3000 - A1)
103: If dx1 < 0 And dy1 < 0 Then Ygolt1 = CInt(3000 + A1)
104: If dx1 > 0 And dy1 < 0 Then Ygolt1 = CInt(6000 - A1)
10411: If Ygolt1 <= 1500 And OH1 >= 4500 Then
      Dovort1 = Ygolt1 + 6000 - OH1
      ElseIf OH1 <= 1500 And Ygolt1 >= 4500 Then
      Dovort1 = Ygolt1 - (OH1 + 6000)
      Else
      Dovort1 = Ygolt1 - OH1
      End If
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = OZ.pZar1
       If zar = "Полн" Then
       v01 = BP.pV01p
       ElseIf zar = "Умен" Then
       v01 = BP.pV01y
       ElseIf zar = "Перв" Then
       v01 = BP.pV011
       ElseIf zar = "Втор" Then
       v01 = BP.pV012
       ElseIf zar = "Трет" Then
       v01 = BP.pV013
       ElseIf zar = "Четверт" Then
       v01 = BP.pV014
       Else
       v01 = BP.pV01p
End If

snar = OZ.pSnar1: vzriv = OZ.pVzr1
OZ.msgVelikaDalnost snar, zar, "1-я Батарея", Dt

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       OZ.poddV0 tz, zar, dv0
              rep1 = OZ.pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
       If rep1 = "Пристрелян" Then
       popvnap = (dDov1 / (Dret1 + 0.001)) * Dt1
       Else
       popvnap = dZwc * Wz + zc
       End If
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If rep1 = "Пристрелян" Then
       popvD = (dDr1 / (Dret1 + 0.001)) * Dt1
       Else
        popvD = dXwc * Wx + dXhc * dhh1 + dXtc * dddt1 + dXv0c * (v01 + dv0)
        Dtk = Dt1 + 1000
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt1 - 1000
            OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt1 + 1000
            OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
                 If popvD < 0 Then
                   Dt = Dt1 - 1000
                   OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt1 + 1000
                   OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   Else
                     popvnapk = dZwc * Wz + zc
                End If
        End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt1 + dXv0c * (v01 + dv0)
       End If
       Dtisch = Dt1 - popvD
       If rep1 = "Пристрелян" Then GoTo 9200
       Dtischk = Dtk - popvdk:
       If popvD < 0 Then
                kPop = (popvD - popvdk) / (Dtisch - Dtischk)
       Else
       kPop = (popvdk - popvD) / (Dtischk - Dtisch)
       End If
       If popvD < 0 Then
       popvD = (Abs(popvD) * kPop - popvD) * -1
       Else
       popvD = Abs(popvD) * kPop + popvD
       End If
       ''''''''''''''''''''''''''''''''''''''''''''''''''
9200:   popvd1 = popvD: Disch = Dt1 + popvD: Disch1 = Disch
                Kpopnap = popvnap - popvnapk
                Kpopnap = Abs(Kpopnap + 0.001) / Abs(Dtisch - Dtischk)
                If popvnap <= 0 And popvnapk >= 0 Then
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                        ElseIf popvnap < 0 And popvnapk <= 0 And popvnapk > popvnap Then
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                        ElseIf popvnap < 0 And popvnapk <= 0 And popvnapk < popvnap Then
                        popvnap = (Kpopnap * Abs(popvD) * -1) + popvnap
                        ElseIf popvnap > 0 And popvnapk > 0 And popvnap > popvnapk Then
                        popvnap = (Kpopnap * Abs(popvD) * -1) + popvnap
                        Else
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                End If
       popvnap1 = popvnap: dovisch1 = Int(Dovort1 + popvnap)
        dhh = dhh1: dddt = dddt1: dV00 = (v01 + dv0): rep = rep1
      If rep1 = "Пристрелян" Then
        dN = (dN / (Dret1 + 0.001)) * Dt1
        Else
      End If
      If snar = "ОФ" And stre = "Мортирная" Then
      OZ.podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      OZ.podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      OZ.pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      OZ.podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      OZ.podT90 Disch, zar, N, dNtus, Pricisch
      Else
      OZ.podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
End If
       If stre = "Мортирная" Then
        Pric1 = Pricisch
        Else
        Pric1 = Pricisch + Yr1
       End If
        Yr = Abs(Yr1): Yrr = Yr1: N1 = N: dNtus1 = dNtus
        If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "РГМ" Then
            Ygpad1 = Ygpad: Ygvozv1 = Ygvozv: Vustra1 = Vustra: ts1 = ts: dXtus11 = dXtus
            Else
            Ygpad1 = Ygpadk: Ygvozv1 = Ygvozvk: Vustra1 = Vustrak: ts1 = tsk: dXtus11 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus1 = 0
       If stre = "Мортирная" Then
       OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr1: preps1 = CInt(Pric1 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr1: preps1 = CInt(Pric1 + daep)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
                Fr = pFrc1: Gl = pGlc1
        veer = Int(Fr / ((Dt1 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX1 <> 0 And pXc1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1
        OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1: OZ.pvVeer1.Text = veer: OZ.pvSk1.Text = Sk
        OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1
        OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
        OZ.pvSnar1.Text = 0: OZ.pvvzr1.Text = 0: OZ.pvZar1.Text = 0: OZ.pvPric1.Text = 0
        OZ.pvN1.Text = 0: OZ.pvDov1.Text = 0: OZ.pvVeer1.Text = 0: OZ.pvSk1.Text = 0
        OZ.pvdXtus1.Text = 0: OZ.pvdNtus1.Text = 0: OZ.pvPolet1.Text = 0: OZ.pvVustra1.Text = 0
        OZ.pvVd1.Text = 0: OZ.pvDt1.Text = 0: OZ.pvYgt1.Text = 0: OZ.pvDovt1.Text = 0
        OZ.pvYr1.Text = 0: OZ.pvOH1.Text = 0: OZ.pvdD1.Text = 0
        OZ.pvDisch1.Text = 0: OZ.pvdDov1.Text = 0
End If
vrv = 0
 ' 2B
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
         Xc2 = pXc2: Yc2 = pYc2: hc2 = phc2
        dx2 = Xc2 - Xop2
104112:  dy2 = Yc2 - Yop2
104113:  dh2 = hc2 - hop2
104114:  Dt2 = Int(Sqr(dx2 ^ 2 + dy2 ^ 2))
104115:  Yr2 = CInt((dh2 / (Dt2 * 0.001 + 0.1)) * 0.95)
104116:  A2 = Abs(Atn(dy2 / (dx2 + 0.001)) / Pi * 30) * 100
104117:  If dx2 > 0 And dy2 > 0 Then Ygolt2 = CInt(A2)
104118:  If dx2 < 0 And dy2 > 0 Then Ygolt2 = CInt(3000 - A2)
104119:  If dx2 < 0 And dy2 < 0 Then Ygolt2 = CInt(3000 + A2)
1041191:  If dx2 > 0 And dy2 < 0 Then Ygolt2 = CInt(6000 - A2)
1041192: If Ygolt2 <= 1500 And OH2 >= 4500 Then
        Dovort2 = Ygolt2 + 6000 - OH2
        ElseIf OH2 <= 1500 And Ygolt2 >= 4500 Then
         Dovort2 = Ygolt2 - (OH2 + 6000)
     Else
         Dovort2 = Ygolt2 - (OH2)
       End If
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = OZ.pZar2
       If zar = "Полн" Then
       v02 = BP.pV02p
       ElseIf zar = "Умен" Then
       v02 = BP.pV02y
       ElseIf zar = "Перв" Then
       v02 = BP.pV021
       ElseIf zar = "Втор" Then
       v02 = BP.pV022
       ElseIf zar = "Трет" Then
       v02 = BP.pV023
       ElseIf zar = "Четверт" Then
       v02 = BP.pV024
       Else
       v02 = BP.pV02p
     End If
     
snar = OZ.pSnar2: vzriv = OZ.pVzr2
OZ.msgVelikaDalnost snar, zar, "2-я Батарея", Dt

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        OZ.poddV0 tz, zar, dv0
               rep2 = OZ.pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
       If rep2 = "Пристрелян" Then
       popvnap = (dDov2 / (Dret2 + 0.001)) * Dt2
       Else
       popvnap = dZwc * Wz + zc
       End If
       dddt2 = dddt
       If rep2 = "Пристрелян" Then
       popvD = (dDr2 / (Dret2 + 0.001)) * Dt2
       Else
       popvD = dXwc * Wx + dXhc * dhh2 + dXtc * dddt2 + dXv0c * (v02 + dv0)
        Dtk = Dt2 + 1000
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt2 - 1000
            OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt2 + 1000
            OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt2 - 1000
                   OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt2 + 1000
                   OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   Else
                     popvnapk = dZwc * Wz + zc
                End If
            End If
        popvdk = dXwc * Wx + dXhc * dhh2 + dXtc * dddt2 + dXv0c * (v02 + dv0)
       End If
       Dtisch = Dt2 - popvD
        If rep2 = "Пристрелян" Then GoTo 9300
       Dtischk = Dtk - popvdk
       If popvD < 0 Then
       kPop = (popvD - popvdk) / (Dtisch - Dtischk)
       Else
       kPop = (popvdk - popvD) / (Dtischk - Dtisch)
       End If
       If popvD < 0 Then
       popvD = (Abs(popvD) * kPop - popvD) * -1
       Else
       popvD = Abs(popvD) * kPop + popvD
       End If
9300:   popvd2 = popvD: Disch = Dt2 + popvD: Disch2 = Disch
                Kpopnap = popvnap - popvnapk
                Kpopnap = Abs(Kpopnap + 0.001) / Abs(Dtisch - Dtischk)
                If popvnap <= 0 And popvnapk >= 0 Then
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                        ElseIf popvnap < 0 And popvnapk <= 0 And popvnapk > popvnap Then
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                        ElseIf popvnap < 0 And popvnapk <= 0 And popvnapk < popvnap Then
                        popvnap = (Kpopnap * Abs(popvD) * -1) + popvnap
                        ElseIf popvnap > 0 And popvnapk > 0 And popvnap > popvnapk Then
                        popvnap = (Kpopnap * Abs(popvD) * -1) + popvnap
                        Else
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                End If
       popvnap2 = popvnap: dovisch2 = CInt(Dovort2 + popvnap)
      dhh = dhh2: dddt = dddt2: dV00 = (v02 + dv0): rep = rep2
      If rep2 = "Пристрелян" Then
        dN = (dN / (Dret2 + 0.001)) * Dt2
        Else
      End If
      If snar = "ОФ" And stre = "Мортирная" Then
      OZ.podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      OZ.podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      OZ.pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      OZ.podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      OZ.podT90 Disch, zar, N, dNtus, Pricisch
      Else
      OZ.podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
End If
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
              Fr = pFrc2: Gl = pGlc2
        veer = Int(Fr / ((Dt2 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX2 <> 0 And pXc2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2
              OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2: OZ.pvVeer2.Text = veer: OZ.pvSk2.Text = Sk
              OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2
        OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD): OZ.pvDisch2.Text = Int(Disch2)
        OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
              OZ.pvSnar2.Text = 0: OZ.pvvzr2.Text = 0: OZ.pvZar2.Text = 0: OZ.pvPric2.Text = 0
              OZ.pvN2.Text = 0: OZ.pvDov2.Text = 0: OZ.pvVeer2.Text = 0: OZ.pvSk2.Text = 0
              OZ.pvdXtus2.Text = 0: OZ.pvdNtus2.Text = 0: OZ.pvPolet2.Text = 0: OZ.pvVustra2.Text = 0
        OZ.pvVd2.Text = 0: OZ.pvDt2.Text = 0: OZ.pvYgt2.Text = 0: OZ.pvDovt2.Text = 0: OZ.pvYr2.Text = 0
        OZ.pvOH2.Text = 0: OZ.pvdD2.Text = 0: OZ.pvDisch2.Text = 0
        OZ.pvdDov2.Text = 0
End If
vrv = 0
  '3B
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
          Xc3 = pXc3: Yc3 = pYc3: hc3 = phc3
         dx3 = Xc3 - Xop3
1041194:  dy3 = Yc3 - Yop3
1041195:  dh3 = hc3 - hop3
1041196:   Dt3 = Int(Sqr(dx3 ^ 2 + dy3 ^ 2))
1041197:   Yr3 = CInt((dh3 / (Dt3 * 0.001 + 0.1)) * 0.95)
1041198:  A3 = Abs(Atn(dy3 / (dx3 + 0.001)) / Pi * 30) * 100
1041199:  If dx3 > 0 And dy3 > 0 Then Ygolt3 = CInt(A3)
10411991:  If dx3 < 0 And dy3 > 0 Then Ygolt3 = CInt(3000 - A3)
10411992:  If dx3 < 0 And dy3 < 0 Then Ygolt3 = CInt(3000 + A3)
10411993:  If dx3 > 0 And dy3 < 0 Then Ygolt3 = CInt(6000 - A3)
10411994:  If Ygolt3 <= 1500 And OH3 >= 4500 Then
          Dovort3 = Ygolt3 + 6000 - OH3
          ElseIf OH3 <= 1500 And Ygolt3 >= 4500 Then
         Dovort3 = Ygolt3 - (OH3 + 6000)
     Else
         Dovort3 = Ygolt3 - (OH3)
       End If
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = OZ.pZar3
       If zar = "Полн" Then
       v03 = BP.pV03p
       ElseIf zar = "Умен" Then
       v03 = BP.pV03Y
       ElseIf zar = "Перв" Then
       v03 = BP.pV031
       ElseIf zar = "Втор" Then
       v03 = BP.pV032
       ElseIf zar = "Трет" Then
       v03 = BP.pV033
       ElseIf zar = "Четверт" Then
       v03 = BP.pV034
       Else
       v03 = BP.pV03p
       End If
       
snar = OZ.pSnar3: vzriv = OZ.pVzr3
OZ.msgVelikaDalnost snar, zar, "3-я Батарея", Dt

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     OZ.poddV0 tz, zar, dv0
            rep3 = OZ.pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
       If rep3 = "Пристрелян" Then
       popvnap = (dDov3 / (Dret3 + 0.001)) * Dt3
       Else
       popvnap = dZwc * Wz + zc
       End If
       dddt3 = dddt
       If rep3 = "Пристрелян" Then
       popvD = (dDr3 / (Dret3 + 0.001)) * Dt3
       Else
       popvD = dXwc * Wx + dXhc * dhh3 + dXtc * dddt3 + dXv0c * (v03 + dv0)
        If q = 35 Then GoTo 9400
        Dtk = Dt3 + 1000
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt3 - 1000
            OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt3 + 1000
            OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt3 - 1000
                   OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt3 + 1000
                   OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   Else
                     popvnapk = dZwc * Wz + zc
                End If
            End If
        popvdk = dXwc * Wx + dXhc * dhh3 + dXtc * dddt3 + dXv0c * (v03 + dv0)
       End If
       Dtisch = Dt3 - popvD
        If rep3 = "Пристрелян" Then GoTo 9400
       Dtischk = Dtk - popvdk
       If popvD < 0 Then
       kPop = (popvD - popvdk) / (Dtisch - Dtischk)
       Else
       kPop = (popvdk - popvD) / (Dtischk - Dtisch)
       End If
       If popvD < 0 Then
       popvD = (Abs(popvD) * kPop - popvD) * -1
       Else
       popvD = Abs(popvD) * kPop + popvD
       End If
9400:   popvd3 = popvD: Disch = Dt3 + popvD: Disch3 = Disch
        If q = 35 Then
                Else
                Kpopnap = popvnap - popvnapk
                Kpopnap = Abs(Kpopnap + 0.001) / Abs(Dtisch - Dtischk)
                If popvnap <= 0 And popvnapk >= 0 Then
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                        ElseIf popvnap < 0 And popvnapk <= 0 And popvnapk > popvnap Then
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                        ElseIf popvnap < 0 And popvnapk <= 0 And popvnapk < popvnap Then
                        popvnap = (Kpopnap * Abs(popvD) * -1) + popvnap
                        ElseIf popvnap > 0 And popvnapk > 0 And popvnap > popvnapk Then
                        popvnap = (Kpopnap * Abs(popvD) * -1) + popvnap
                        Else
                        popvnap = Kpopnap * Abs(popvD) + popvnap
                End If
       End If
       popvnap3 = popvnap: dovisch3 = CInt(Dovort3 + popvnap)
      dhh = dhh3: dddt = dddt3: dV00 = (v03 + dv0): rep = rep3
      If rep3 = "Пристрелян" Then
        dN = (dN / (Dret3 + 0.001)) * Dt3
        Else
      End If
       If snar = "ОФ" And stre = "Мортирная" Then
      OZ.podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      OZ.podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      OZ.pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      OZ.podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      OZ.podT90 Disch, zar, N, dNtus, Pricisch
      Else
      OZ.podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
      End If
       If stre = "Мортирная" Then
        Pric3 = Pricisch
        Else
        Pric3 = Pricisch + Yr3
       End If
        Yr = Abs(Yr3): Yrr = Yr3: N3 = N: dNtus3 = dNtus
If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "РГМ" Then
            Ygpad3 = Ygpad: Ygvozv3 = Ygvozv: Vustra3 = Vustra: ts3 = ts: dXtus3 = dXtus
            Else
            Ygpad3 = Ygpadk: Ygvozv3 = Ygvozvk: Vustra3 = Vustrak: ts3 = tsk: dXtus3 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus3 = 0
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
               Fr = pFrc3: Gl = pGlc3
        veer = Int(Fr / ((Dt3 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX3 <> 0 And pXc3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3
                     OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3: OZ.pvVeer3.Text = veer: OZ.pvSk3.Text = Sk
                     OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3
        OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD): OZ.pvDisch3.Text = Int(Disch3)
        OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
                     OZ.pvSnar3.Text = 0: OZ.pvvzr3.Text = 0: OZ.pvZar3.Text = 0: OZ.pvPric3.Text = 0
                     OZ.pvN3.Text = 0: OZ.pvDov3.Text = 0: OZ.pvVeer3.Text = 0: OZ.pvSk3.Text = 0
                     OZ.pvdXtus3.Text = 0: OZ.pvdNtus3.Text = 0: OZ.pvPolet3.Text = 0: OZ.pvVustra3.Text = 0
        OZ.pvVd3.Text = 0: OZ.pvDt3.Text = 0: OZ.pvYgt3.Text = 0: OZ.pvDovt3.Text = 0: OZ.pvYr3.Text = 0
        OZ.pvOH3.Text = 0: OZ.pvdD3.Text = 0: OZ.pvDisch3.Text = 0
        OZ.pvdDov3.Text = 0
End If
vrv = 0

End Sub

Private Sub Command2_Click()
OZzelkagdform.Hide
End Sub
Private Sub pvvNc1_Click()
Dim nz As String
Dim xc As Single, yc As Single, hc As Single, Fr As Single, Gl As Single
nz = pvvNc1
1011 Open "D:\YO_NA\Zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, z1, z2, z3, z4, z5, z6
   If z1 = nz Then xc = z2: yc = z3: hc = z4: Fr = z5: Gl = z6: GoTo 1012
        GoTo 101111
1012 Close #1
pXc1.Text = xc: pYc1.Text = yc: phc1.Text = hc: pFrc1.Text = Fr: pGlc1.Text = Gl
End Sub
Private Sub pvvNc1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim z(1 To 10) As String
Dim nz As String
Dim xc As Single, yc As Single, hc As Single
nz = pvvNc1
If KeyCode = 13 Then
1011    Open "D:\YO_NA\zeli" For Input As #1
101111  If EOF(1) Then GoTo 1012
   Input #1, z(1), z(2), z(3), z(4), z(5), z(6)
   If z(1) = nz Then xc = z(2): yc = z(3): hc = z(4): Fr = z(5): Gl = z(6): GoTo 1012
        GoTo 101111
1012    Close #1
pXc1.Text = xc: pYc1.Text = yc: phc1.Text = hc: pFrc1.Text = Fr: pGlc1.Text = Gl
    Else
End If
End Sub

Private Sub pvvNc2_Click()
Dim nz As String
Dim xc As Single, yc As Single, hc As Single, Fr As Single, Gl As Single
nz = pvvNc2
1011 Open "D:\YO_NA\Zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, z1, z2, z3, z4, z5, z6
   If z1 = nz Then xc = z2: yc = z3: hc = z4: Fr = z5: Gl = z6: GoTo 1012
        GoTo 101111
1012 Close #1
pXc2.Text = xc: pYc2.Text = yc: phc2.Text = hc: pFrc2.Text = Fr: pGlc2.Text = Gl
End Sub
Private Sub pvvNc2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim z(1 To 10) As String
Dim nz As String
Dim xc As Single, yc As Single, hc As Single
nz = pvvNc2
If KeyCode = 13 Then
1011    Open "D:\YO_NA\zeli" For Input As #1
101111  If EOF(1) Then GoTo 1012
   Input #1, z(1), z(2), z(3), z(4), z(5), z(6)
   If z(1) = nz Then xc = z(2): yc = z(3): hc = z(4): Fr = z(5): Gl = z(6): GoTo 1012
        GoTo 101111
1012    Close #1
pXc2.Text = xc: pYc2.Text = yc: phc2.Text = hc: pFrc2.Text = Fr: pGlc2.Text = Gl
    Else
End If
End Sub

Private Sub pvvNc3_Click()
Dim nz As String
Dim xc As Single, yc As Single, hc As Single, Fr As Single, Gl As Single
nz = pvvNc3
1011 Open "D:\YO_NA\Zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, z1, z2, z3, z4, z5, z6
   If z1 = nz Then xc = z2: yc = z3: hc = z4: Fr = z5: Gl = z6: GoTo 1012
        GoTo 101111
1012 Close #1
pXc3.Text = xc: pYc3.Text = yc: phc3.Text = hc: pFrc3.Text = Fr: pGlc3.Text = Gl
End Sub
Private Sub pvvNc3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim z(1 To 10) As String
Dim nz As String
Dim xc As Single, yc As Single, hc As Single
nz = pvvNc3
If KeyCode = 13 Then
1011    Open "D:\YO_NA\zeli" For Input As #1
101111  If EOF(1) Then GoTo 1012
   Input #1, z(1), z(2), z(3), z(4), z(5), z(6)
   If z(1) = nz Then xc = z(2): yc = z(3): hc = z(4): Fr = z(5): Gl = z(6): GoTo 1012
        GoTo 101111
1012    Close #1
pXc3.Text = xc: pYc3.Text = yc: phc3.Text = hc: pFrc3.Text = Fr: pGlc3.Text = Gl
    Else
End If
End Sub

Private Sub pXc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc1.Text = ""
pYc1.SetFocus
End If
End Sub
Private Sub pYc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc1.Text = ""
phc1.SetFocus
End If
End Sub
Private Sub phc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pFrc1.Text = ""
pFrc1.SetFocus
End If
End Sub
Private Sub pFrc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pGlc1.Text = ""
pGlc1.SetFocus
End If
End Sub
Private Sub pXc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc2.Text = ""
pYc2.SetFocus
End If
End Sub
Private Sub pYc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc2.Text = ""
phc2.SetFocus
End If
End Sub
Private Sub phc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pFrc2.Text = ""
pFrc2.SetFocus
End If
End Sub
Private Sub pFrc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pGlc2.Text = ""
pGlc2.SetFocus
End If
End Sub
Private Sub pXc3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc3.Text = ""
pYc3.SetFocus
End If
End Sub
Private Sub pYc3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc3.Text = ""
phc3.SetFocus
End If
End Sub
Private Sub phc3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pFrc3.Text = ""
pFrc3.SetFocus
End If
End Sub
Private Sub pFrc3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pGlc3.Text = ""
pGlc3.SetFocus
End If
End Sub
Private Sub Form_Load()
Dim t(1 To 10) As String
941 Open "D:\YO_NA\zeli" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 942
 Input #1, t(1), t(2), t(3), t(4), t(5), t(6)
pvvNc1.AddItem t(1)
Loop
942 Close #1
9412 Open "D:\YO_NA\zeli" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 9422
 Input #1, t(1), t(2), t(3), t(4), t(5), t(6)
pvvNc2.AddItem t(1)
Loop
9422 Close #1
9413 Open "D:\YO_NA\zeli" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 9423
 Input #1, t(1), t(2), t(3), t(4), t(5), t(6)
pvvNc3.AddItem t(1)
Loop
9423 Close #1
End Sub
