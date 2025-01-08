VERSION 5.00
Begin VB.Form otprkomandy 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Отправить команду"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optEach 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option2"
      Height          =   200
      Left            =   4000
      TabIndex        =   15
      Top             =   900
      Width           =   300
   End
   Begin VB.OptionButton optEveryone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option1"
      Height          =   200
      Left            =   900
      TabIndex        =   14
      Top             =   900
      Value           =   -1  'True
      Width           =   300
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
      ItemData        =   "otprkomandy.frx":0000
      Left            =   300
      List            =   "otprkomandy.frx":0013
      TabIndex        =   11
      Text            =   "Осколочный"
      Top             =   4800
      Width           =   5000
   End
   Begin VB.CommandButton otprVFail 
      BackColor       =   &H00FF8080&
      Caption         =   "Отправить в файл"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   2100
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
      ItemData        =   "otprkomandy.frx":006B
      Left            =   300
      List            =   "otprkomandy.frx":007E
      TabIndex        =   5
      Text            =   "Записать"
      Top             =   6600
      Width           =   5000
   End
   Begin VB.CheckBox pBat2 
      BackColor       =   &H00C0C0C0&
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
      Left            =   1800
      TabIndex        =   3
      Top             =   3400
      Width           =   375
   End
   Begin VB.CheckBox pBat3 
      BackColor       =   &H00C0C0C0&
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
      Left            =   3300
      TabIndex        =   2
      Top             =   3400
      Width           =   735
   End
   Begin VB.CheckBox pBat1 
      BackColor       =   &H00C0C0C0&
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
      Left            =   400
      TabIndex        =   1
      Top             =   3400
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Цель каждому"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3000
      TabIndex        =   13
      Top             =   300
      Width           =   2200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Цель всем"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   300
      TabIndex        =   12
      Top             =   300
      Width           =   1600
   End
   Begin VB.Label Label7 
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
      Height          =   495
      Left            =   700
      TabIndex        =   10
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Left            =   2900
      TabIndex        =   8
      Top             =   2700
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   1500
      TabIndex        =   7
      Top             =   2700
      Width           =   900
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   100
      TabIndex        =   6
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Исполнительная команда"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   300
      TabIndex        =   4
      Top             =   5600
      Width           =   3135
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
      Height          =   500
      Left            =   700
      TabIndex        =   0
      Top             =   2040
      Width           =   2000
   End
End
Attribute VB_Name = "otprkomandy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub otprVFail_Click()
Dim nZeli As String, nOr As String, snar As String, vzriv As String, zar As String, Pric As String, dovorot As String, ispKom As String, trybka As Single
Dim dovIsch As Single
Dim trybkaZapis As String

Komanda.Show

If optEveryone = True Then
    Open App.Path & "\numberZeli" For Input As #1
    Input #1, nZeli
    nZeli1 = nZeli: nZeli2 = nZeli: nZeli3 = nZeli
    Close #1
    Else
    Open App.Path & "\writeZeliEach" For Input As #1
    Input #1, nZeli1, nZeli2, nZeli3
    Close #1
End If

If pBat1 = 1 Then
    nOr = "Десна": snar = OZ.pvSnar1:  zar = OZ.pvZar1: Pric = OZ.pvPric1: dovIsch = OZ.pvDov1: ispKom = pIspKom: trybka = OZ.pvN1
    
    If snar = "3Ш" Then
        vzriv = OZ.pvvzr1
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = OZ.pvvzr1
            If vzriv = "РГМ" Then vzriv = pVzriv
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = OZ.pvvzr1
            vzriv = vzriv + ", " + str(trybka)
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komanda.txt" For Output As #1
    Write #1, nOr + ", Цель " + nZeli1 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom
    Close #1
    Komanda.pvKomand.Text = nOr + ", Цель " + nZeli1 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + "__" & vbCrLf
    Else
    Open App.Path & "\komanda.txt" For Output As #1
    Komanda.pvKomand.Text = ""
    Close #1
    End If
    
If pBat2 = 1 Then
    nOr = "Самара": snar = OZ.pvSnar2:  zar = OZ.pvZar2: Pric = OZ.pvPric2: dovIsch = OZ.pvDov2: ispKom = pIspKom: trybka = OZ.pvN2
    
    If snar = "3Ш" Then
        vzriv = OZ.pvvzr2
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = OZ.pvvzr2
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = OZ.pvvzr2
            vzriv = vzriv + ", " + str(trybka)
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komanda.txt" For Append As #1
    Write #1, nOr + ", Цель " + nZeli2 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom
    Close #1
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli2 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + "__" & vbCrLf
    Else
    End If
    
If pBat3 = 1 Then
    nOr = "Корень": snar = OZ.pvSnar3:  zar = OZ.pvZar3: Pric = OZ.pvPric3: dovIsch = OZ.pvDov3: ispKom = pIspKom: trybka = OZ.pvN3
    
    If snar = "3Ш" Then
        vzriv = OZ.pvvzr3
        trybkaZapis = " Трубка " + CStr(trybka)
        ElseIf snar = "ОФ" Then
            vzriv = OZ.pvvzr3
            If vzriv = "АР-5" Or vzriv = "В-90" Then
                vzriv = vzriv + ", " + str(trybka)
                Else
            End If
        ElseIf snar = "С4" Then
            vzriv = OZ.pvvzr3
            vzriv = vzriv + ", " + str(trybka)
        Else
        vzriv = pVzriv
        trybkaZapis = ""
    End If
        
    If dovIsch >= 0 Then
        dovorot = "OH +" + CStr(dovIsch)
        Else
        dovorot = "OH " + CStr(dovIsch)
        End If
    Open App.Path & "\komanda.txt" For Append As #1
    Write #1, nOr + ", Цель " + nZeli3 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom
    Close #1
    Komanda.pvKomand.Text = Komanda.pvKomand.Text & nOr + ", Цель " + nZeli3 + ", Снаряд " + snar + ", Взрыватель " + vzriv + ", Заряд " + zar + ", Прицел " + Pric + trybkaZapis + ", Доворот " + dovorot + ". " + ispKom + "__" & vbCrLf
    Else
    End If
    
End Sub
