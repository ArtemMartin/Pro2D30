VERSION 5.00
Begin VB.Form Pristrelka 
   BackColor       =   &H0000C0C0&
   Caption         =   "Пристрелка"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13260
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   12
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Сопряженка"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      Left            =   4440
      TabIndex        =   34
      Top             =   3200
      Width           =   9000
      Begin VB.CommandButton Soprkor 
         BackColor       =   &H008080FF&
         Caption         =   "Решить"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1500
         Width           =   1200
      End
      Begin VB.TextBox pMcp 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5100
         TabIndex        =   44
         Text            =   "0"
         Top             =   2520
         Width           =   1700
      End
      Begin VB.TextBox pAp 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5100
         TabIndex        =   43
         Text            =   "0"
         Top             =   1500
         Width           =   1700
      End
      Begin VB.TextBox pMcl 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1300
         TabIndex        =   42
         Text            =   "0"
         Top             =   2500
         Width           =   1700
      End
      Begin VB.TextBox pAl 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1300
         TabIndex        =   41
         Text            =   "0"
         Top             =   1500
         Width           =   1700
      End
      Begin VB.Label Label20 
         BackColor       =   &H0000C0C0&
         Caption         =   "Мц="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   4300
         TabIndex        =   40
         Top             =   2500
         Width           =   600
      End
      Begin VB.Label Label19 
         BackColor       =   &H0000C0C0&
         Caption         =   "А="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   4300
         TabIndex        =   39
         Top             =   1500
         Width           =   500
      End
      Begin VB.Label Label18 
         BackColor       =   &H0000C0C0&
         Caption         =   "Мц="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   500
         TabIndex        =   38
         Top             =   2500
         Width           =   600
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000C0C0&
         Caption         =   "А="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   500
         TabIndex        =   37
         Top             =   1500
         Width           =   500
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000C0C0&
         Caption         =   "           Правый"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4300
         TabIndex        =   36
         Top             =   480
         Width           =   2505
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000C0C0&
         Caption         =   "            Левый"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   500
         TabIndex        =   35
         Top             =   480
         Width           =   2500
      End
   End
   Begin VB.CommandButton XY 
      BackColor       =   &H008080FF&
      Caption         =   "Решить"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   16350
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1200
      Width           =   1200
   End
   Begin VB.TextBox pY 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   14350
      TabIndex        =   32
      Text            =   "0"
      Top             =   2200
      Width           =   1700
   End
   Begin VB.TextBox pX 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   14350
      TabIndex        =   31
      Text            =   "0"
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton dXdY 
      BackColor       =   &H008080FF&
      Caption         =   "Решить"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1200
      Width           =   1200
   End
   Begin VB.TextBox pdY 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9900
      TabIndex        =   26
      Text            =   "0"
      Top             =   2200
      Width           =   1700
   End
   Begin VB.TextBox pdX 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9900
      TabIndex        =   25
      Text            =   "0"
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton NZR 
      BackColor       =   &H008080FF&
      Caption         =   "Решить"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   7400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1200
      Width           =   1200
   End
   Begin VB.TextBox pdD 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5500
      TabIndex        =   20
      Text            =   "0"
      Top             =   2200
      Width           =   1700
   End
   Begin VB.TextBox pdA 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5500
      TabIndex        =   19
      Text            =   "0"
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton DAK 
      BackColor       =   &H008080FF&
      Caption         =   "Решить"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1200
      Width           =   1200
   End
   Begin VB.TextBox pkorYr 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      TabIndex        =   14
      Text            =   "0"
      Top             =   7200
      Width           =   1700
   End
   Begin VB.TextBox pkorYg 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      TabIndex        =   13
      Text            =   "0"
      Top             =   6200
      Width           =   1700
   End
   Begin VB.TextBox pkorD 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      TabIndex        =   12
      Text            =   "0"
      Top             =   5200
      Width           =   1700
   End
   Begin VB.TextBox pMcraz 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      TabIndex        =   11
      Text            =   "0"
      Top             =   3200
      Width           =   1700
   End
   Begin VB.TextBox pDraz 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      TabIndex        =   10
      Text            =   "0"
      Top             =   2200
      Width           =   1700
   End
   Begin VB.TextBox pAraz 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton vuxod 
      BackColor       =   &H00808080&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8000
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackColor       =   &H0000C0C0&
      Caption         =   "Yр="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   13600
      TabIndex        =   30
      Top             =   2200
      Width           =   500
   End
   Begin VB.Label Label22 
      BackColor       =   &H0000C0C0&
      Caption         =   "Xр="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   13600
      TabIndex        =   29
      Top             =   1200
      Width           =   500
   End
   Begin VB.Label Label21 
      BackColor       =   &H0000C0C0&
      Caption         =   "                X, Y"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   13600
      TabIndex        =   28
      Top             =   350
      Width           =   2400
   End
   Begin VB.Label Label17 
      BackColor       =   &H0000C0C0&
      Caption         =   "dY="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9200
      TabIndex        =   24
      Top             =   2200
      Width           =   500
   End
   Begin VB.Label Label16 
      BackColor       =   &H0000C0C0&
      Caption         =   "dX="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9200
      TabIndex        =   23
      Top             =   1200
      Width           =   500
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000C0C0&
      Caption         =   "              dX, dY"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9200
      TabIndex        =   22
      Top             =   350
      Width           =   2400
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Caption         =   "dД="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4700
      TabIndex        =   18
      Top             =   2200
      Width           =   600
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "dA="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4800
      TabIndex        =   17
      Top             =   1200
      Width           =   500
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "                НЗР"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4800
      TabIndex        =   16
      Top             =   350
      Width           =   2400
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ур="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   480
      TabIndex        =   8
      Top             =   7200
      Width           =   600
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "Дов="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   480
      TabIndex        =   7
      Top             =   6200
      Width           =   700
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "Д="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   480
      TabIndex        =   6
      Top             =   5200
      Width           =   400
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "    Корректура"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   500
      TabIndex        =   5
      Top             =   4200
      Width           =   2400
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Мцр="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   500
      TabIndex        =   4
      Top             =   3200
      Width           =   700
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Др="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   500
      TabIndex        =   3
      Top             =   2200
      Width           =   500
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ар="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   500
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "                   ДАК"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   500
      TabIndex        =   1
      Top             =   350
      Width           =   2400
   End
End
Attribute VB_Name = "Pristrelka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Xnp As Single
Public Ynp As Single
Public hnp As Single
Public Xop As Single
Public Yop As Single
Public hop As Single
Dim dx As Single
Dim dy As Single
Public Dt As Single
Public Yr As Single
Public Pi, Ygolt As Single
Dim Araz As Single
Dim Draz As Single
Dim Mcraz As Single
Dim dxr, dyr, Arr, Xraz, Yraz, hr, korD, korYg, korYr, Yrr As Single
Dim dA, Ac, Dc, dD, Dr As Single

Private Sub DAK_Click()
Xop = topo.pXop: Yop = topo.pYop: hop = topo.phop
Xnp = topo.pXnp: Ynp = topo.pYnp: hnp = topo.phnp
Dt = topo.pDt: Ygolt = topo.pYgt: Yr = topo.pYr
Araz = pAraz: Draz = pDraz: Mcr = pMcraz
Xraz = Cos((Araz + 0.001) / 100 * 6 * 3.141592 / 180) * Draz + Xnp
Yraz = Sin((Araz + 0.001) / 100 * 6 * 3.141592 / 180) * Draz + Ynp
hraz = Mcr * ((Draz + 0.001) / 1000) * 1.05 + hnp
Pi = 3.14159265358
dxr = Xraz - Xop: dyr = Yraz - Yop: dhraz = hraz - hop
Dtr = Sqr(dxr ^ 2 + dyr ^ 2)
Yrr = (dhraz / (Dtr * 0.001 + 0.001)) * 0.95
 Arr = Abs(Atn(dyr / (dxr + 0.001)) / Pi * 30) * 100
If dxr > 0 And dyr > 0 Then Ygoltr = Arr
If dxr < 0 And dyr > 0 Then Ygoltr = 3000 - Arr
If dxr < 0 And dyr < 0 Then Ygoltr = 3000 + Arr
If dxr > 0 And dyr < 0 Then Ygoltr = 6000 - Arr
korD = Dt - Dtr: korYr = Yr - Yrr
If Ygolt < 1500 And Ygoltr > 4500 Then
    korYg = Ygolt + 6000 - Ygoltr
    Else
    korYg = Ygolt - Ygoltr
End If
pkorD.Text = Format(korD, "0")
pkorYg.Text = Format(korYg, "0")
pkorYr.Text = Format(korYr, "0")
End Sub

Private Sub dXdY_Click()
Dt = topo.pDt: Ygolt = topo.pYgt
Xop = topo.pXop: Yop = topo.pYop: hop = topo.phop
Xz = topo.pXz: Yz = topo.pYz
dx = pdX: dy = pdY
Xraz = Xz + dx: Yraz = Yz + dy
Pi = 3.14159265358
dxr = Xraz - Xop: dyr = Yraz - Yop: dhraz = hraz - hop
Dtr = Sqr(dxr ^ 2 + dyr ^ 2)
 Arr = Abs(Atn(dyr / (dxr + 0.001)) / Pi * 30) * 100
If dxr > 0 And dyr > 0 Then Ygoltr = Arr
If dxr < 0 And dyr > 0 Then Ygoltr = 3000 - Arr
If dxr < 0 And dyr < 0 Then Ygoltr = 3000 + Arr
If dxr > 0 And dyr < 0 Then Ygoltr = 6000 - Arr
korD = Dt - Dtr
If Ygolt < 1500 And Ygoltr > 4500 Then
    korYg = Ygolt + 6000 - Ygoltr
    Else
    korYg = Ygolt - Ygoltr
End If
pkorD.Text = Format(korD, "0")
pkorYg.Text = Format(korYg, "0")
End Sub

Private Sub NZR_Click()
Ac = topo.A: Dc = topo.D
dA = pdA: dD = pdD
Araz = Ac + dA: Draz = Dc + dD
Xop = topo.pXop: Yop = topo.pYop: hop = topo.phop
Xnp = topo.pXnp: Ynp = topo.pYnp: hnp = topo.phnp
Dt = topo.pDt: Ygolt = topo.pYgt: Yr = topo.pYr
Xraz = Cos((Araz + 0.001) / 100 * 6 * 3.141592 / 180) * Draz + Xnp
Yraz = Sin((Araz + 0.001) / 100 * 6 * 3.141592 / 180) * Draz + Ynp
Pi = 3.14159265358
dxr = Xraz - Xop: dyr = Yraz - Yop: dhraz = hraz - hop
Dtr = Sqr(dxr ^ 2 + dyr ^ 2)
Arr = Abs(Atn(dyr / (dxr + 0.001)) / Pi * 30) * 100
If dxr > 0 And dyr > 0 Then Ygoltr = Arr
If dxr < 0 And dyr > 0 Then Ygoltr = 3000 - Arr
If dxr < 0 And dyr < 0 Then Ygoltr = 3000 + Arr
If dxr > 0 And dyr < 0 Then Ygoltr = 6000 - Arr
korD = Dt - Dtr
If Ygolt < 1500 And Ygoltr > 4500 Then
    korYg = Ygolt + 6000 - Ygoltr
    Else
    korYg = Ygolt - Ygoltr
End If
pkorD.Text = Format(korD, "0")
pkorYg.Text = Format(korYg, "0")
End Sub

Private Sub Soprkor_Click()
Xop = topo.pXop: Yop = topo.pYop: hop = topo.phop
Xnp = topo.pXnp: Ynp = topo.pYnp: hnp = topo.phnp
Dt = topo.pDt: Ygolt = topo.pYgt: Yr = topo.pYr
Xp = topo.pXp: Yp = topo.pYp: hp = topo.php
Xl = topo.pXl: Yl = topo.pYl: hl = topo.phl
Alev = pAl: Aprav = pAp
Mcp = pMcp: Mcl = pMcl
dxso = Xp - Xl: dyso = Yp - Yl
  baz = Sqr(dxso ^ 2 + dyso ^ 2)
  aso = Abs(Atn(dyso / (dxso + 0.1)) / 3.141592 * 30) * 100
  If dxso > 0 And dyso > 0 Then Ygolbaz = Int(aso)
  If dxso < 0 And dyso > 0 Then Ygolbaz = Int(3000 - aso)
  If dxso < 0 And dyso < 0 Then Ygolbaz = Int(3000 + aso)
  If dxso > 0 And dyso < 0 Then Ygolbaz = Int(6000 - aso)
  If Alev < 1500 And Aprav > 4500 Then fi = Abs(Alev + 6000 - Aprav)
  If Alev > 4500 And Aprav < 1500 Then fi = Abs(Alev - (Aprav + 6000))
  If Alev > Aprav Then fi = Abs(Alev - Aprav)
  If Alev < 1500 And Ygolbaz > 4500 Then
        blev = Abs(Alev + 6000 - Ygolbaz)
      ElseIf Alev > 4500 And Ygolbaz < 1500 Then
        blev = Abs(Alev - (Ygo.baz + 6000))
      Else
        blev = Abs(Alev - Ygolbaz)
  End If
  If Ygolbaz - 3000 < 0 Then
  ybazp = Ygolbaz + 3000
  Else
  ybazp = Ygolbaz - 3000
  End If
  If Aprav < 1500 And ybazp > 4500 Then
  bprav = Abs(Aprav + 6000 - ybazp)
  ElseIf Aprav > 4500 And ybazp < 1500 Then
  bprav = Abs(Aprav - (ybazp + 6000))
  Else
  bprav = Abs(Aprav - ybazp)
  End If
  Dlev = Abs(baz / (Sin(fi / 100 * 6 * 3.141592 / 180) + 0.001) * Sin(bprav / 100 * 6 * 3.141592 / 180))
  Dprav = Abs(baz / (Sin(fi / 100 * 6 * 3.141592 / 180) + 0.001) * Sin(blev / 100 * 6 * 3.141592 / 180))
  Xraz = Cos(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Xl
  Yraz = Sin(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Yl
  If Mcl = 0 Then hraz = Mcp * (Dprav * 0.001) * 1.05 + hp
  If Mcp = 0 Then hraz = Mcl * (Dlev * 0.001) * 1.05 + hl
  Pi = 3.14159265358
dxr = Xraz - Xop: dyr = Yraz - Yop: dhraz = hraz - hop
Dtr = Sqr(dxr ^ 2 + dyr ^ 2)
Yrr = (dhraz / (Dtr * 0.001 + 0.001)) * 0.95
 Arr = Abs(Atn(dyr / (dxr + 0.001)) / Pi * 30) * 100
If dxr > 0 And dyr > 0 Then Ygoltr = Arr
If dxr < 0 And dyr > 0 Then Ygoltr = 3000 - Arr
If dxr < 0 And dyr < 0 Then Ygoltr = 3000 + Arr
If dxr > 0 And dyr < 0 Then Ygoltr = 6000 - Arr
korD = Dt - Dtr: korYr = Yr - Yrr
If Ygolt < 1500 And Ygoltr > 4500 Then
    korYg = Ygolt + 6000 - Ygoltr
    Else
    korYg = Ygolt - Ygoltr
End If
pkorD.Text = Format(korD, "0")
pkorYg.Text = Format(korYg, "0")
pkorYr.Text = Format(korYr, "0")
End Sub

Private Sub vuxod_Click()
Unload Pristrelka
End Sub

Private Sub XY_Click()
Dt = topo.pDt: Ygolt = topo.pYgt
Xop = topo.pXop: Yop = topo.pYop
Xraz = pX: Yraz = pY
Pi = 3.14159265358
dxr = Xraz - Xop: dyr = Yraz - Yop: dhraz = hraz - hop
Dtr = Sqr(dxr ^ 2 + dyr ^ 2)
 Arr = Abs(Atn(dyr / (dxr + 0.001)) / Pi * 30) * 100
If dxr > 0 And dyr > 0 Then Ygoltr = Arr
If dxr < 0 And dyr > 0 Then Ygoltr = 3000 - Arr
If dxr < 0 And dyr < 0 Then Ygoltr = 3000 + Arr
If dxr > 0 And dyr < 0 Then Ygoltr = 6000 - Arr
korD = Dt - Dtr
If Ygolt < 1500 And Ygoltr > 4500 Then
    korYg = Ygolt + 6000 - Ygoltr
    Else
    korYg = Ygolt - Ygoltr
End If
pkorD.Text = Format(korD, "0")
pkorYg.Text = Format(korYg, "0")
End Sub
