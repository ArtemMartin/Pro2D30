VERSION 5.00
Begin VB.Form KontrolPooryd 
   Caption         =   "Êîíòðîëü ïîîðóäèéíî"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18645
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   14.25
      Charset         =   204
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   18645
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "ÂÛÕÎÄ"
      Height          =   1200
      Left            =   16450
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5300
      Width           =   2000
   End
   Begin VB.CommandButton reshit 
      BackColor       =   &H00FF8080&
      Caption         =   "ÐÅØÈÒÜ"
      Height          =   1200
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5300
      Width           =   2000
   End
   Begin VB.CommandButton prinDan 
      BackColor       =   &H00FF8080&
      Caption         =   "Ïðèíÿòü äàííûå ñ ÎÇ"
      Height          =   1200
      Left            =   100
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5300
      Width           =   2000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   5000
      Left            =   12500
      TabIndex        =   36
      Top             =   100
      Width           =   6000
      Begin VB.TextBox osk3IschZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00004080&
         Height          =   500
         Left            =   4500
         TabIndex        =   74
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox osk3IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00004080&
         Height          =   500
         Left            =   4500
         TabIndex        =   73
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox osk3Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00004080&
         Height          =   500
         Left            =   4500
         TabIndex        =   72
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox osk3ZTN 
         ForeColor       =   &H00004080&
         Height          =   500
         Left            =   4500
         TabIndex        =   71
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox osk3OTN 
         ForeColor       =   &H00004080&
         Height          =   500
         Left            =   4500
         TabIndex        =   70
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.TextBox osk2IschZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00800080&
         Height          =   500
         Left            =   3000
         TabIndex        =   53
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox osk2IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00800080&
         Height          =   500
         Left            =   3000
         TabIndex        =   52
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox osk2Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00800080&
         Height          =   500
         Left            =   3000
         TabIndex        =   51
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox osk2ZTN 
         ForeColor       =   &H00800080&
         Height          =   500
         Left            =   3000
         TabIndex        =   50
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox osk2OTN 
         ForeColor       =   &H00800080&
         Height          =   500
         Left            =   3000
         TabIndex        =   49
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.TextBox osk1IschZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1500
         TabIndex        =   48
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox osk1IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1500
         TabIndex        =   47
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox osk1Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1500
         TabIndex        =   46
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox osk1ZTN 
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1500
         TabIndex        =   45
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox osk1OTN 
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1500
         TabIndex        =   44
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.Label labOr9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Êîð3"
         ForeColor       =   &H00004080&
         Height          =   500
         Left            =   4680
         TabIndex        =   69
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Èñ÷ ïî ÇÒÍ"
         Height          =   700
         Left            =   100
         TabIndex        =   43
         Top             =   3900
         Width           =   1200
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Èñ÷ ïî ÎÒÍ"
         Height          =   700
         Left            =   100
         TabIndex        =   42
         Top             =   3100
         Width           =   1200
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Äîâ ÎÍ"
         Height          =   500
         Left            =   100
         TabIndex        =   41
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÇÒÍ"
         Height          =   500
         Left            =   100
         TabIndex        =   40
         Top             =   1700
         Width           =   800
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÎÒÍ"
         Height          =   500
         Left            =   100
         TabIndex        =   39
         Top             =   1000
         Width           =   800
      End
      Begin VB.Label labOr8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Êîð2"
         ForeColor       =   &H00800080&
         Height          =   500
         Left            =   3200
         TabIndex        =   38
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labOr7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Êîð1"
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1680
         TabIndex        =   37
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   5000
      Left            =   6300
      TabIndex        =   18
      Top             =   100
      Width           =   6000
      Begin VB.TextBox kal3IschZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   500
         Left            =   4500
         TabIndex        =   68
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox kal3IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   500
         Left            =   4500
         TabIndex        =   67
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox kal3Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   500
         Left            =   4500
         TabIndex        =   66
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox kal3ZTN 
         ForeColor       =   &H00008000&
         Height          =   500
         Left            =   4500
         TabIndex        =   65
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox kal3OTN 
         ForeColor       =   &H00008000&
         Height          =   500
         Left            =   4500
         TabIndex        =   64
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.TextBox kal2IschZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008080&
         Height          =   500
         Left            =   3000
         TabIndex        =   35
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox kal2IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008080&
         Height          =   500
         Left            =   3000
         TabIndex        =   34
         Text            =   "0"
         Top             =   3120
         Width           =   1200
      End
      Begin VB.TextBox kal2Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008080&
         Height          =   500
         Left            =   3000
         TabIndex        =   33
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox kal2ZTN 
         ForeColor       =   &H00008080&
         Height          =   500
         Left            =   3000
         TabIndex        =   32
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox kal2OTN 
         ForeColor       =   &H00008080&
         Height          =   500
         Left            =   3000
         TabIndex        =   31
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.TextBox kal1IschZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H000040C0&
         Height          =   500
         Left            =   1500
         TabIndex        =   30
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox kal1IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H000040C0&
         Height          =   500
         Left            =   1500
         TabIndex        =   29
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox kal1Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H000040C0&
         Height          =   500
         Left            =   1500
         TabIndex        =   28
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox kal1ZTN 
         ForeColor       =   &H000040C0&
         Height          =   500
         Left            =   1500
         TabIndex        =   27
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox kal1OTN 
         ForeColor       =   &H000040C0&
         Height          =   500
         Left            =   1500
         TabIndex        =   26
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.Label labOr6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ñàì3"
         ForeColor       =   &H00008000&
         Height          =   500
         Left            =   4700
         TabIndex        =   63
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Èñ÷ ïî ÇÒÍ"
         Height          =   700
         Left            =   100
         TabIndex        =   25
         Top             =   3900
         Width           =   1200
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Èñ÷ ïî ÎÒÍ"
         Height          =   700
         Left            =   100
         TabIndex        =   24
         Top             =   3100
         Width           =   1200
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Äîâ ÎÍ"
         Height          =   500
         Left            =   100
         TabIndex        =   23
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÇÒÍ"
         Height          =   500
         Left            =   100
         TabIndex        =   22
         Top             =   1700
         Width           =   800
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÎÒÍ"
         Height          =   500
         Left            =   100
         TabIndex        =   21
         Top             =   1000
         Width           =   800
      End
      Begin VB.Label labOr5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ñàì2"
         ForeColor       =   &H00008080&
         Height          =   500
         Left            =   3200
         TabIndex        =   20
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labOr4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ñàì1"
         ForeColor       =   &H000040C0&
         Height          =   500
         Left            =   1700
         TabIndex        =   19
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Frame KontPoor 
      BackColor       =   &H00C0C0C0&
      Height          =   5000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   6000
      Begin VB.TextBox aks3OTN 
         ForeColor       =   &H00C000C0&
         Height          =   500
         Left            =   4500
         TabIndex        =   62
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.TextBox aks3ZTN 
         ForeColor       =   &H00C000C0&
         Height          =   500
         Left            =   4500
         TabIndex        =   61
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox aks3Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C000C0&
         Height          =   500
         Left            =   4500
         TabIndex        =   60
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox aks3IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C000C0&
         Height          =   500
         Left            =   4500
         TabIndex        =   59
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox aks3IschZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C000C0&
         Height          =   500
         Left            =   4500
         TabIndex        =   58
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox aks2IschZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H000000FF&
         Height          =   500
         Left            =   3000
         TabIndex        =   17
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox aks2IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H000000FF&
         Height          =   500
         Left            =   3000
         TabIndex        =   16
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox aks2Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H000000FF&
         Height          =   500
         Left            =   3000
         TabIndex        =   15
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox aks2ZTN 
         ForeColor       =   &H000000FF&
         Height          =   500
         Left            =   3000
         TabIndex        =   14
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox aks2OTN 
         ForeColor       =   &H000000FF&
         Height          =   500
         Left            =   3000
         TabIndex        =   13
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.TextBox aks1IshcZTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   500
         Left            =   1500
         TabIndex        =   11
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox aks1IschOTN 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   500
         Left            =   1500
         TabIndex        =   10
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox aks1Dov 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   500
         Left            =   1500
         TabIndex        =   9
         Text            =   "0"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.TextBox aks1ZTN 
         Height          =   500
         Left            =   1500
         TabIndex        =   8
         Text            =   "0"
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox aks1OTN 
         Height          =   500
         Left            =   1500
         TabIndex        =   7
         Text            =   "0"
         Top             =   1000
         Width           =   1200
      End
      Begin VB.Label labOr3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Äåñ3"
         ForeColor       =   &H00C000C0&
         Height          =   500
         Left            =   4700
         TabIndex        =   57
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labOr2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Äåñ2"
         ForeColor       =   &H000000FF&
         Height          =   500
         Left            =   3300
         TabIndex        =   12
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Èñ÷ ïî ÇÒÍ"
         Height          =   700
         Left            =   240
         TabIndex        =   6
         Top             =   3900
         Width           =   1200
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Èñ÷ ïî ÎÒÍ"
         Height          =   700
         Left            =   240
         TabIndex        =   5
         Top             =   3100
         Width           =   1200
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Äîâ ÎÍ"
         Height          =   500
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÇÒÍ"
         Height          =   500
         Left            =   240
         TabIndex        =   3
         Top             =   1700
         Width           =   700
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÎÒÍ"
         Height          =   500
         Left            =   240
         TabIndex        =   2
         Top             =   1000
         Width           =   700
      End
      Begin VB.Label labOr1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Äåñ1"
         Height          =   500
         Left            =   1680
         TabIndex        =   1
         Top             =   300
         Width           =   1200
      End
   End
End
Attribute VB_Name = "KontrolPooryd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
KontrolPooryd.Hide
End Sub

Private Sub Form_Load()
labOr1 = Shest6Oryd.labOr1
labOr2 = Shest6Oryd.labOr2
labOr3 = Shest6Oryd.labOr3
labOr4 = Shest6Oryd.labOr4
labOr5 = Shest6Oryd.labOr5
labOr6 = Shest6Oryd.labOr6
labOr7 = Shest6Oryd.labOr7
labOr8 = Shest6Oryd.labOr8
labOr9 = Shest6Oryd.labOr9
End Sub

Private Sub prinDan_Click()
aks1Dov.Text = Shest6Oryd.pvDov1
aks2Dov.Text = Shest6Oryd.pvDov2
aks3Dov.Text = Shest6Oryd.pvDov3
kal1Dov.Text = Shest6Oryd.pvDov4
kal2Dov.Text = Shest6Oryd.pvDov5
kal3Dov.Text = Shest6Oryd.pvDov6
osk1Dov.Text = Shest6Oryd.pvDov7
osk2Dov.Text = Shest6Oryd.pvDov8
osk3Dov.Text = Shest6Oryd.pvDov9
End Sub

Private Sub reshit_Click()
Dim perAks1OTN As Single, perAks1ZTN As Single
Dim perAks2OTN As Single, perAks2ZTN As Single
Dim perAks3OTN As Single, perAks3ZTN As Single

Dim perKal1OTN As Single, perKal1ZTN As Single
Dim perKal2OTN As Single, perKal2ZTN As Single
Dim perKal3OTN As Single, perKal3ZTN As Single

Dim perOsk1OTN As Single, perOsk1ZTN As Single
Dim perOsk2OTN As Single, perOsk2ZTN As Single
Dim perOsk3OTN As Single, perOsk3ZTN As Single

Dim perAsk1Dov As Single, perAsk2Dov As Single, perAsk3Dov As Single
Dim perKal1Dov As Single, perKal2Dov As Single, perKal3Dov As Single
Dim perOsk1Dov As Single, perOsk2Dov As Single, perOsk3Dov As Single

Dim perAks1IschOTN As Single, perAks1IschZTN As Single
Dim perAks2IschOTN As Single, perAks2IschZTN As Single
Dim perAks3IschOTN As Single, perAks3IschZTN As Single

Dim perKal1IschOTN As Single, perKal1IschZTN As Single
Dim perKal2IschOTN As Single, perKal2IschZTN As Single
Dim perKal3IschOTN As Single, perKal3IschZTN As Single

Dim perOsk1IschOTN As Single, perOsk1IschZTN As Single
Dim perOsk2IschOTN As Single, perOsk2IschZTN As Single
Dim perOsk3IschOTN As Single, perOsk3IschZTN As Single

perAks1OTN = aks1OTN: perAks1ZTN = aks1ZTN
perAks2OTN = aks2OTN: perAks2ZTN = aks2ZTN
perAks3OTN = aks3OTN: perAks3ZTN = aks3ZTN

perKal1OTN = kal1OTN: perKal1ZTN = kal1ZTN
perKal2OTN = kal2OTN: perKal2ZTN = kal2ZTN
perKal3OTN = kal3OTN: perKal3ZTN = kal3ZTN

perOsk1OTN = osk1OTN: perOsk1ZTN = osk1ZTN
perOsk2OTN = osk2OTN: perOsk2ZTN = osk2ZTN
perOsk3OTN = osk3OTN: perOsk3ZTN = osk3ZTN

perAsk1Dov = aks1Dov: perAsk2Dov = aks2Dov: perAsk3Dov = aks3Dov
perKal1Dov = kal1Dov: perKal2Dov = kal2Dov: perKal3Dov = kal3Dov
perOsk1Dov = osk1Dov: perOsk2Dov = osk2Dov: perOsk3Dov = osk3Dov
'Àêñàé
If perAks1OTN + perAsk1Dov < 0 Then
    perAks1IschOTN = perAks1OTN + perAsk1Dov + 6000
    Else
    perAks1IschOTN = perAks1OTN + perAsk1Dov
End If
If perAks1ZTN + perAsk1Dov < 0 Then
    perAks1IschZTN = perAks1ZTN + perAsk1Dov + 6000
    Else
    perAks1IschZTN = perAks1ZTN + perAsk1Dov
End If

If perAks2OTN + perAsk2Dov < 0 Then
    perAks2IschOTN = perAks2OTN + perAsk2Dov + 6000
    Else
    perAks2IschOTN = perAks2OTN + perAsk2Dov
End If
If perAks2ZTN + perAsk2Dov < 0 Then
    perAks2IschZTN = perAks2ZTN + perAsk2Dov + 6000
    Else
    perAks2IschZTN = perAks2ZTN + perAsk2Dov
End If

If perAks3OTN + perAsk3Dov < 0 Then
    perAks3IschOTN = perAks3OTN + perAsk3Dov + 6000
    Else
    perAks3IschOTN = perAks3OTN + perAsk3Dov
End If
If perAks3ZTN + perAsk3Dov < 0 Then
    perAks3IschZTN = perAks3ZTN + perAsk3Dov + 6000
    Else
    perAks3IschZTN = perAks3ZTN + perAsk3Dov
End If

If perAsk1Dov = 0 Then
    perAks1IschOTN = 0
    perAks1IschZTN = 0
    Else
End If
If perAsk2Dov = 0 Then
    perAks2IschOTN = 0
    perAks2IschZTN = 0
    Else
End If
If perAsk3Dov = 0 Then
    perAks3IschOTN = 0
    perAks3IschZTN = 0
    Else
End If

'Êàëèòâà
If perKal1OTN + perKal1Dov < 0 Then
    perKal1IschOTN = perKal1OTN + perKal1Dov + 6000
    Else
    perKal1IschOTN = perKal1OTN + perKal1Dov
End If
If perKal1ZTN + perKal1Dov < 0 Then
    perKal1IschZTN = perKal1ZTN + perKal1Dov + 6000
    Else
    perKal1IschZTN = perKal1ZTN + perKal1Dov
End If

If perKal2OTN + perKal2Dov < 0 Then
    perKal2IschOTN = perKal2OTN + perKal2Dov + 6000
    Else
    perKal2IschOTN = perKal2OTN + perKal2Dov
End If
If perKal2ZTN + perKal2Dov < 0 Then
    perKal2IschZTN = perKal2ZTN + perKal2Dov + 6000
    Else
    perKal2IschZTN = perKal2ZTN + perKal2Dov
End If

If perKal3OTN + perKal3Dov < 0 Then
    perKal3IschOTN = perKal3OTN + perKal3Dov + 6000
    Else
    perKal3IschOTN = perKal3OTN + perKal3Dov
End If
If perKal3ZTN + perKal3Dov < 0 Then
    perKal3IschZTN = perKal3ZTN + perKal3Dov + 6000
    Else
    perKal3IschZTN = perKal3ZTN + perKal3Dov
End If

If perKal1Dov = 0 Then
    perKal1IschOTN = 0
    perKal1IschZTN = 0
    Else
End If
If perKal2Dov = 0 Then
    perKal2IschOTN = 0
    perKal2IschZTN = 0
    Else
End If
If perKal3Dov = 0 Then
    perKal3IschOTN = 0
    perKal3IschZTN = 0
    Else
End If

'Îñêîë
If perOsk1OTN + perOsk1Dov < 0 Then
    perOsk1IschOTN = perOsk1OTN + perOsk1Dov + 6000
    Else
    perOsk1IschOTN = perOsk1OTN + perOsk1Dov
End If
If perOsk1ZTN + perOsk1Dov < 0 Then
    perOsk1IschZTN = perOsk1ZTN + perOsk1Dov + 6000
    Else
    perOsk1IschZTN = perOsk1ZTN + perOsk1Dov
End If

If perOsk2OTN + perOsk2Dov < 0 Then
    perOsk2IschOTN = perOsk2OTN + perOsk2Dov + 6000
    Else
    perOsk2IschOTN = perOsk2OTN + perOsk2Dov
End If
If perOsk2ZTN + perOsk2Dov < 0 Then
    perOsk2IschZTN = perOsk2ZTN + perOsk2Dov + 6000
    Else
    perOsk2IschZTN = perOsk2ZTN + perOsk2Dov
End If

If perOsk3OTN + perOsk3Dov < 0 Then
    perOsk3IschOTN = perOsk3OTN + perOsk3Dov + 6000
    Else
    perOsk3IschOTN = perOsk3OTN + perOsk3Dov
End If
If perOsk3ZTN + perOsk3Dov < 0 Then
    perOsk3IschZTN = perOsk3ZTN + perOsk3Dov + 6000
    Else
    perOsk3IschZTN = perOsk3ZTN + perOsk3Dov
End If

If perOsk1Dov = 0 Then
    perOsk1IschOTN = 0
    perOsk1IschZTN = 0
    Else
End If
If perOsk2Dov = 0 Then
    perOsk2IschOTN = 0
    perOsk2IschZTN = 0
    Else
End If
If perOsk3Dov = 0 Then
    perOsk3IschOTN = 0
    perOsk3IschZTN = 0
    Else
End If

aks1IschOTN.Text = Round(perAks1IschOTN)
aks1IshcZTN.Text = Round(perAks1IschZTN)
aks2IschOTN.Text = Round(perAks2IschOTN)
aks2IschZTN.Text = Round(perAks2IschZTN)
aks3IschOTN.Text = Round(perAks3IschOTN)
aks3IschZTN.Text = Round(perAks3IschZTN)

kal1IschOTN.Text = Round(perKal1IschOTN)
kal1IschZTN.Text = Round(perKal1IschZTN)
kal2IschOTN.Text = Round(perKal2IschOTN)
kal2IschZTN.Text = Round(perKal2IschZTN)
kal3IschOTN.Text = Round(perKal3IschOTN)
kal3IschZTN.Text = Round(perKal3IschZTN)

osk1IschOTN.Text = Round(perOsk1IschOTN)
osk1IschZTN.Text = Round(perOsk1IschZTN)
osk2IschOTN.Text = Round(perOsk2IschOTN)
osk2IschZTN.Text = Round(perOsk2IschZTN)
osk3IschOTN.Text = Round(perOsk3IschOTN)
osk3IschZTN.Text = Round(perOsk3IschZTN)

End Sub
