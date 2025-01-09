VERSION 5.00
Begin VB.MDIForm OsnYO 
   BackColor       =   &H8000000C&
   Caption         =   "Управление огнем Д30"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13785
   LinkTopic       =   "MDIForm1"
   Picture         =   "OsnYO.frx":0000
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   13725
      TabIndex        =   0
      Top             =   0
      Width           =   13785
      Begin VB.CommandButton Command8 
         Caption         =   "ПЗО"
         Enabled         =   0   'False
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
         Left            =   11400
         TabIndex        =   8
         Top             =   100
         Width           =   2000
      End
      Begin VB.CommandButton Command7 
         Caption         =   "9 Орудий"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   17700
         TabIndex        =   7
         Top             =   120
         Width           =   1500
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0000FFFF&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   19300
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   100
         Width           =   800
      End
      Begin VB.CommandButton Command5 
         Caption         =   "НЗО"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9300
         TabIndex        =   5
         Top             =   100
         Width           =   2000
      End
      Begin VB.CommandButton Command4 
         Caption         =   "РЕПЕР"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7200
         TabIndex        =   4
         Top             =   100
         Width           =   2000
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Боевой порядок"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   3
         Top             =   100
         Width           =   2500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Огневая задача"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2700
         TabIndex        =   2
         Top             =   100
         Width           =   2300
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Пристрелка"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5100
         TabIndex        =   1
         Top             =   100
         Width           =   2000
      End
   End
End
Attribute VB_Name = "OsnYO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Xop1, Yop1, hop1, OH1 As Single
Public Xop2, Yop2, hop2, OH2 As Single
Public Xop3, Yop3, hop3, OH3 As Single
Public Xnp1, Ynp1, hnp1 As Single
Public Xnp2, Ynp2, hnp2 As Single
Public Xnp3, Ynp3, hnp3 As Single
Public Xnp4, Ynp4, hnp4 As Single
Public Xnp5, Ynp5, hnp5 As Single
Public tz1, tz2, tz3, V0P1, V0P2, V0P3 As Single
Public V0Y1, V0Y2, V0Y3, V011, V012, V013, V021, V022, V023, V031, V032, V033, V041, V042, V043 As Single
Public hmet, h, Tv, Aw, W As Single
Public dT02, dT04, dT08, dT12, dT16, dT20, dT24, dT30, dT40, dT50, dT60 As Single
Public Aw02, Aw04, Aw08, Aw12, Aw16, Aw20, Aw24, Aw30, Aw40Aw50, Aw60 As Single
Public W02, W04, W08, W12, W16, W20, W24, W30, W40, W50, W60 As Single
Public Snar1, Snar2, Snar3 As String
Public Vzriv1, Vzriv2, Vzriv3, zar1, zar2, zar3 As String
Public Xc, Yc, hc, Frc, Glc, Ac, Dc, Mc, Alev, Aprav, Mclev, Mcprav As Single
Public Pric1, Pric2, Pric3, N1, N2, N3, Dovor1, Dovor2, Dovor3, Veer1, Veer2, Veer3, Sk1, Sk2, Sk3 As Single
Public dXtus1, dXtus2, dXtus3, dNtus1, dNtus2, dNtus3, ts1, ts2, ts3, Vustr1, Vustr2, Vustr3, Vd1, Vd2, Vd3, Dt1, Dt2, Dt3 As Single
Public Ygolt1, Ygolt2, Ygolt3, Dovort1, Dovort2, Dovort3, Yr1, Yr2, Yr3, dD1, dD2, dD3, Disch1, Disch2, Disch3, dDov1, dDv2, dDov3 As Single


Private Sub Command1_Click()
Pristrelka.Show
End Sub

Private Sub Command2_Click()
OZ.Show
End Sub

Private Sub Command3_Click()
BP.Show
End Sub

Private Sub Command4_Click()
REPER.Show
End Sub

Private Sub Command5_Click()
NZO.Show
End Sub

Private Sub Command6_Click()
Razrabotchik.Show
End Sub

Private Sub Command7_Click()
Shest6Oryd.Show
End Sub

Private Sub Command8_Click()
PZO.Show
End Sub

Private Sub MDIForm_Load()
BP.Show
BP.Hide
REPER.Show
REPER.Hide
End Sub
