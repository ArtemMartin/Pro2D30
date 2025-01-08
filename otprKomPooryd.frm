VERSION 5.00
Begin VB.Form otprKomPooryd 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Отправка команды поорудийно"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   8.25
      Charset         =   204
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   15300
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optEach 
      BackColor       =   &H00C0C0C0&
      Height          =   400
      Left            =   4000
      TabIndex        =   16
      Top             =   800
      Width           =   300
   End
   Begin VB.OptionButton optEveryone 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   204
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   800
      TabIndex        =   15
      Top             =   800
      Value           =   -1  'True
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Записать в файл"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   13000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   400
      Width           =   1935
   End
   Begin VB.ComboBox pIspKom 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "otprKomPooryd.frx":0000
      Left            =   300
      List            =   "otprKomPooryd.frx":0013
      TabIndex        =   13
      Text            =   "Записать"
      Top             =   6700
      Width           =   6000
   End
   Begin VB.ComboBox pVzriv 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "otprKomPooryd.frx":005F
      Left            =   300
      List            =   "otprKomPooryd.frx":0072
      TabIndex        =   11
      Text            =   "Осколочный"
      Top             =   5100
      Width           =   6000
   End
   Begin VB.CheckBox pChOsk3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Кор3"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   360
      Left            =   12800
      TabIndex        =   9
      Top             =   4200
      Width           =   2200
   End
   Begin VB.CheckBox pChOsk2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Кор2"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   10300
      TabIndex        =   8
      Top             =   4200
      Width           =   2200
   End
   Begin VB.CheckBox pChOsk1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Кор1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Left            =   7800
      TabIndex        =   7
      Top             =   4200
      Width           =   2200
   End
   Begin VB.CheckBox pChKal3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Сам3"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   8900
      TabIndex        =   6
      Top             =   3600
      Width           =   2200
   End
   Begin VB.CheckBox pChKal2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Сам2"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   360
      Left            =   6500
      TabIndex        =   5
      Top             =   3600
      Width           =   2200
   End
   Begin VB.CheckBox pChKal1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Сам1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   4000
      TabIndex        =   4
      Top             =   3600
      Width           =   2200
   End
   Begin VB.CheckBox pChAks3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Дес3"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Left            =   5200
      TabIndex        =   3
      Top             =   3000
      Width           =   2200
   End
   Begin VB.CheckBox pChAks2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Дес2"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   2800
      TabIndex        =   2
      Top             =   3000
      Width           =   2200
   End
   Begin VB.CheckBox pChAks1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Дес1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   1
      Top             =   3000
      Width           =   2200
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Цель каждому"
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
      Left            =   3000
      TabIndex        =   18
      Top             =   300
      Width           =   2500
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Цель всем"
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
      Left            =   300
      TabIndex        =   17
      Top             =   300
      Width           =   1700
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Исполнительная"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   12
      Top             =   6000
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Взрыватель"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   10
      Top             =   4400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Привлечь"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "otprKomPooryd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim nZeli As String, nOr As String, snar As String, vzriv As String, zar As String, Pric As String, dovorot As String, ispKom As String, trybka As Single
Dim dovIsch As Single
Dim trybkaZapis As String
Dim mesage As String
Dim nZeli1 As String, nZeli2 As String, nZeli3 As String, nZeli4 As String
Dim nZeli5 As String, nZeli6 As String, nZeli7 As String, nZeli8 As String
Dim nZeli9 As String

Komanda.Show

If optEveryone = True Then
    Open App.Path & "\numberZeli" For Input As #1
    Input #1, nZeli
    nZeli1 = nZeli: nZeli2 = nZeli: nZeli3 = nZeli: nZeli4 = nZeli
    nZeli5 = nZeli: nZeli6 = nZeli: nZeli7 = nZeli: nZeli8 = nZeli
    nZeli9 = nZeli
    Close #1
    Else
    Open App.Path & "\writeZeliEach" For Input As #1
    Input #1, nZeli1, nZeli2, nZeli3, nZeli4, nZeli5, nZeli6, nZeli7, nZeli8, nZeli9
    Close #1
End If

If pChAks1 = 1 Then
    nOr = pChAks1.Caption: snar = Shest6Oryd.pAks1Snar:  zar = Shest6Oryd.pAks1Zar: Pric = Shest6Oryd.pvPric1: dovIsch = Shest6Oryd.pvDov1
    ispKom = pIspKom: trybka = Shest6Oryd.pvN1
    
    If snar = "3Ш" Then
        vzriv = Shest6Oryd.pAks1Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pAks1Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pAks1Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch > 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Output As #1
    Write #1, nOr + ", Цель " + nZeli1 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Close #1
    Komanda.pvKomand.Text = nOr + ", Цель " + nZeli1 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Else
    Open App.Path & "\komPooryd.txt" For Output As #1
    Close #1
    Komanda.pvKomand.Text = ""
    End If
    
If pChAks2 = 1 Then
    nOr = pChAks2.Caption: snar = Shest6Oryd.pAks2Snar:  zar = Shest6Oryd.pAks2Zar: Pric = Shest6Oryd.pvPric2: dovIsch = Shest6Oryd.pvDov2
    ispKom = pIspKom: trybka = Shest6Oryd.pvN2
    
    If snar = "3Ш" Then
        vzriv = Shest6Oryd.pAks2Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pAks2Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pAks2Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Append As #1
    Write #1, nOr + ", Цель " + nZeli2 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Close #1
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli2 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Else
    End If
    
If pChAks3 = 1 Then
    nOr = pChAks3.Caption: snar = Shest6Oryd.pAks3Snar:  zar = Shest6Oryd.pAks3Zar: Pric = Shest6Oryd.pvPric3: dovIsch = Shest6Oryd.pvDov3
    ispKom = pIspKom: trybka = Shest6Oryd.pvN3
    
    If snar = "3Ш" Then
        vzriv = Shest6Oryd.pAks3Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pAks3Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pAks3Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Append As #1
    Write #1, nOr + ", Цель " + nZeli3 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Close #1
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli3 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Else
    End If

If pChKal1 = 1 Then
    nOr = pChKal1.Caption: snar = Shest6Oryd.pKal1Snar:  zar = Shest6Oryd.pKal1Zar: Pric = Shest6Oryd.pvPric4: dovIsch = Shest6Oryd.pvDov4
    ispKom = pIspKom: trybka = Shest6Oryd.pvN4
    
    If snar = "3Ш" Then
        vzriv = Shest6Oryd.pKal1Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pKal1Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pKal1Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Append As #1
    Write #1, nOr + ", Цель " + nZeli4 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Close #1
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli4 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Else
    End If

If pChKal2 = 1 Then
    nOr = pChKal2.Caption: snar = Shest6Oryd.pKal2Snar:  zar = Shest6Oryd.pKal2Zar: Pric = Shest6Oryd.pvPric5: dovIsch = Shest6Oryd.pvDov5
    ispKom = pIspKom: trybka = Shest6Oryd.pvN5
    
    If snar = "3Ш" Then
        vzriv = Shest6Oryd.pKal2Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pKal2Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pKal2Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Append As #1
    Write #1, nOr + ", Цель " + nZeli5 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli5 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Close #1
    Else
    End If

If pChKal3 = 1 Then
    nOr = pChKal3.Caption: snar = Shest6Oryd.pKal3Snar:  zar = Shest6Oryd.pKal3Zar: Pric = Shest6Oryd.pvPric6: dovIsch = Shest6Oryd.pvDov6
    ispKom = pIspKom: trybka = Shest6Oryd.pvN6
    
    If snar = "3Ш" Then
        vzriv = Shest6Oryd.pKal3Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pKal3Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pKal3Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Append As #1
    Write #1, nOr + ", Цель " + nZeli6 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli6 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Close #1
    Else
    End If

If pChOsk1 = 1 Then
    nOr = pChOsk1.Caption: snar = Shest6Oryd.pOsk1Snar:  zar = Shest6Oryd.pOsk1Zar: Pric = Shest6Oryd.pvPric7: dovIsch = Shest6Oryd.pvDov7
    ispKom = pIspKom: trybka = Shest6Oryd.pvN7
    
     If snar = "3Ш" Then
        vzriv = Shest6Oryd.pOsk1Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pOsk1Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pOsk1Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Append As #1
    Write #1, nOr + ", Цель " + nZeli7 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli7 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Close #1
    Else
    End If

If pChOsk2 = 1 Then
    nOr = pChOsk2.Caption: snar = Shest6Oryd.pOsk2Snar:  zar = Shest6Oryd.pOsk2Zar: Pric = Shest6Oryd.pvPric8: dovIsch = Shest6Oryd.pvDov8
    ispKom = pIspKom: trybka = Shest6Oryd.pvN8
    
    If snar = "3Ш" Then
        vzriv = Shest6Oryd.pOsk2Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pOsk2Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pOsk2Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Append As #1
    Write #1, nOr + ", Цель " + nZeli8 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli8 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Close #1
    Else
    End If

If pChOsk3 = 1 Then
    nOr = pChOsk3.Caption: snar = Shest6Oryd.pOsk3Snar:  zar = Shest6Oryd.pOsk3Zar: Pric = Shest6Oryd.pvPric9: dovIsch = Shest6Oryd.pvDov9
    ispKom = pIspKom: trybka = Shest6Oryd.pvN9
    
   If snar = "3Ш" Then
        vzriv = Shest6Oryd.pOsk3Vzr
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = Shest6Oryd.pOsk3Vzr
            If vzriv = "РГМ" Then vzriv = pVzriv: trybkaZapis = ""
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                trybkaZapis = ""
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = Shest6Oryd.pOsk3Vzr
            vzriv = vzriv + ", " + str(trybka)
            trybkaZapis = ""
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komPooryd.txt " For Append As #1
    Write #1, nOr + ", Цель " + nZeli9 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". "
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli9 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + ". " + "__" & vbCrLf
    Close #1
    Else
    End If

End Sub

Private Sub Form_Load()
pChAks1.Caption = Shest6Oryd.labOr1
pChAks2.Caption = Shest6Oryd.labOr2
pChAks3.Caption = Shest6Oryd.labOr3

pChKal1.Caption = Shest6Oryd.labOr4
pChKal2.Caption = Shest6Oryd.labOr5
pChKal3.Caption = Shest6Oryd.labOr6

pChOsk1.Caption = Shest6Oryd.labOr7
pChOsk2.Caption = Shest6Oryd.labOr8
pChOsk3.Caption = Shest6Oryd.labOr9

End Sub
