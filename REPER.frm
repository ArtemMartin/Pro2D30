VERSION 5.00
Begin VB.Form REPER 
   Caption         =   "Репер"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   14.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   18150
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   6050
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3 Батарея пристрелянные поправки"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1900
      Left            =   13200
      TabIndex        =   104
      Top             =   7500
      Width           =   6400
      Begin VB.TextBox ptime3 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         TabIndex        =   122
         Text            =   "0"
         Top             =   1440
         Width           =   2200
      End
      Begin VB.TextBox pvdn3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3800
         TabIndex        =   118
         Text            =   "0"
         Top             =   1000
         Width           =   600
      End
      Begin VB.TextBox pvdDov3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2350
         TabIndex        =   116
         Text            =   "0"
         Top             =   1000
         Width           =   800
      End
      Begin VB.TextBox pvdD3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   600
         TabIndex        =   114
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvDr3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5150
         TabIndex        =   112
         Text            =   "0"
         Top             =   400
         Width           =   1000
      End
      Begin VB.TextBox pvvzr3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3800
         TabIndex        =   110
         Text            =   "0"
         Top             =   400
         Width           =   700
      End
      Begin VB.TextBox pvsn3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2300
         TabIndex        =   108
         Text            =   "0"
         Top             =   400
         Width           =   800
      End
      Begin VB.TextBox pvZar3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   700
         TabIndex        =   106
         Text            =   "0"
         Top             =   400
         Width           =   1000
      End
      Begin VB.Label Label59 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Время созд. РЕПЕРА"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   125
         Top             =   1450
         Width           =   2000
      End
      Begin VB.Label Label56 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dN="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3300
         TabIndex        =   117
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label Label55 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dДов="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1700
         TabIndex        =   115
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label54 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   113
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Др.т."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4600
         TabIndex        =   111
         Top             =   400
         Width           =   500
      End
      Begin VB.Label Label38 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Взрыв"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3200
         TabIndex        =   109
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Снар"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   107
         Top             =   400
         Width           =   500
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Заряд"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   105
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2 Батарея пристрелянные поправки"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1900
      Left            =   6650
      TabIndex        =   89
      Top             =   7500
      Width           =   6400
      Begin VB.TextBox ptime2 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         TabIndex        =   121
         Text            =   "0"
         Top             =   1450
         Width           =   2200
      End
      Begin VB.TextBox pvdn2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3800
         TabIndex        =   103
         Text            =   "0"
         Top             =   1000
         Width           =   600
      End
      Begin VB.TextBox pvdDov2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2350
         TabIndex        =   101
         Text            =   "0"
         Top             =   1000
         Width           =   800
      End
      Begin VB.TextBox pvdD2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   600
         TabIndex        =   99
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvDr2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5150
         TabIndex        =   97
         Text            =   "0"
         Top             =   400
         Width           =   1000
      End
      Begin VB.TextBox pvvzr2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3800
         TabIndex        =   95
         Text            =   "0"
         Top             =   400
         Width           =   700
      End
      Begin VB.TextBox pvsn2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2300
         TabIndex        =   93
         Text            =   "0"
         Top             =   400
         Width           =   800
      End
      Begin VB.TextBox pvZar2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   700
         TabIndex        =   91
         Text            =   "0"
         Top             =   400
         Width           =   1000
      End
      Begin VB.Label Label58 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Время созд. РЕПЕРА"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   124
         Top             =   1450
         Width           =   2000
      End
      Begin VB.Label Label53 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dN="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3300
         TabIndex        =   102
         Top             =   1000
         Width           =   375
      End
      Begin VB.Label Label52 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dДов="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1700
         TabIndex        =   100
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label51 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   98
         Top             =   1000
         Width           =   375
      End
      Begin VB.Label Label50 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Др.т."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4600
         TabIndex        =   96
         Top             =   400
         Width           =   500
      End
      Begin VB.Label Label49 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Взрыв"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3200
         TabIndex        =   94
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Снар"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   92
         Top             =   400
         Width           =   450
      End
      Begin VB.Label Label47 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Заряд"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   90
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1 Батарея пристрелянные поправки"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1900
      Left            =   100
      TabIndex        =   74
      Top             =   7500
      Width           =   6400
      Begin VB.TextBox ptime1 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         TabIndex        =   120
         Text            =   "0"
         Top             =   1450
         Width           =   2200
      End
      Begin VB.TextBox pvDr1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5150
         TabIndex        =   88
         Text            =   "0"
         Top             =   400
         Width           =   1000
      End
      Begin VB.TextBox pvdn1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3800
         TabIndex        =   86
         Text            =   "0"
         Top             =   1000
         Width           =   600
      End
      Begin VB.TextBox pvdDov1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2350
         TabIndex        =   84
         Text            =   "0"
         Top             =   1000
         Width           =   800
      End
      Begin VB.TextBox pvdD1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   600
         TabIndex        =   82
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvvzr1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3800
         TabIndex        =   80
         Text            =   "0"
         Top             =   400
         Width           =   700
      End
      Begin VB.TextBox pvsn1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2300
         TabIndex        =   78
         Text            =   "0"
         Top             =   400
         Width           =   800
      End
      Begin VB.TextBox pvZar1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   700
         TabIndex        =   76
         Text            =   "0"
         Top             =   400
         Width           =   1000
      End
      Begin VB.Label Label57 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Время созд. РЕПЕРА"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   123
         Top             =   1450
         Width           =   2000
      End
      Begin VB.Label Label46 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Др.т."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4600
         TabIndex        =   87
         Top             =   400
         Width           =   495
      End
      Begin VB.Label Label45 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dN="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3300
         TabIndex        =   85
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dДов="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1700
         TabIndex        =   83
         Top             =   1000
         Width           =   615
      End
      Begin VB.Label Label43 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   81
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label Label42 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Взрыв"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3200
         TabIndex        =   79
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label41 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Снар"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   77
         Top             =   400
         Width           =   500
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Заряд"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   75
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Засечка Х, У"
      Height          =   4100
      Left            =   13300
      TabIndex        =   62
      Top             =   3300
      Width           =   4300
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Решить"
         Height          =   1000
         Left            =   2600
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1400
         Width           =   1200
      End
      Begin VB.TextBox phre 
         Height          =   450
         Left            =   700
         TabIndex        =   68
         Text            =   "0"
         Top             =   3000
         Width           =   1000
      End
      Begin VB.TextBox pYre 
         Height          =   450
         Left            =   700
         TabIndex        =   67
         Text            =   "0"
         Top             =   2200
         Width           =   1500
      End
      Begin VB.TextBox pXre 
         Height          =   450
         Left            =   700
         TabIndex        =   66
         Text            =   "0"
         Top             =   1400
         Width           =   1500
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
         Caption         =   "         Координаты Репера"
         Height          =   400
         Left            =   100
         TabIndex        =   70
         Top             =   600
         Width           =   3700
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   400
         Left            =   100
         TabIndex        =   65
         Top             =   3000
         Width           =   500
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   100
         TabIndex        =   64
         Top             =   2200
         Width           =   500
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Х="
         Height          =   400
         Left            =   100
         TabIndex        =   63
         Top             =   1400
         Width           =   500
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Засечка Сопряженным"
      Height          =   4100
      Left            =   5100
      TabIndex        =   46
      Top             =   3300
      Width           =   8000
      Begin VB.CommandButton RepAA 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Решить"
         Height          =   1000
         Left            =   6400
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   1400
         Width           =   1200
      End
      Begin VB.TextBox pMp 
         Height          =   450
         Left            =   4400
         TabIndex        =   60
         Text            =   "0"
         Top             =   3000
         Width           =   1000
      End
      Begin VB.TextBox pAp 
         Height          =   450
         Left            =   4400
         TabIndex        =   59
         Text            =   "0"
         Top             =   2200
         Width           =   1500
      End
      Begin VB.TextBox pMl 
         Height          =   450
         Left            =   1000
         TabIndex        =   56
         Text            =   "0"
         Top             =   3000
         Width           =   1000
      End
      Begin VB.TextBox pAl 
         Height          =   450
         Left            =   1000
         TabIndex        =   55
         Text            =   "0"
         Top             =   2200
         Width           =   1500
      End
      Begin VB.ComboBox pNKPP 
         Height          =   450
         ItemData        =   "REPER.frx":0000
         Left            =   4400
         List            =   "REPER.frx":000D
         TabIndex        =   52
         Text            =   "1"
         Top             =   1400
         Width           =   1000
      End
      Begin VB.ComboBox pNKPL 
         Height          =   450
         ItemData        =   "REPER.frx":001A
         Left            =   1000
         List            =   "REPER.frx":002D
         TabIndex        =   50
         Text            =   "1"
         Top             =   1400
         Width           =   1000
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мц="
         Height          =   400
         Left            =   3500
         TabIndex        =   58
         Top             =   3000
         Width           =   600
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         Height          =   400
         Left            =   3500
         TabIndex        =   57
         Top             =   2200
         Width           =   500
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мц="
         Height          =   400
         Left            =   100
         TabIndex        =   54
         Top             =   3000
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         Height          =   400
         Left            =   100
         TabIndex        =   53
         Top             =   2200
         Width           =   500
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ КП"
         Height          =   400
         Left            =   3500
         TabIndex        =   51
         Top             =   1400
         Width           =   800
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ КП"
         Height          =   400
         Left            =   100
         TabIndex        =   49
         Top             =   1400
         Width           =   800
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         Caption         =   "         ПРАВЫЙ"
         Height          =   400
         Left            =   3500
         TabIndex        =   48
         Top             =   600
         Width           =   2500
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "           ЛЕВЫЙ"
         Height          =   400
         Left            =   100
         TabIndex        =   47
         Top             =   600
         Width           =   2500
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Засечка А, Д"
      Height          =   4100
      Left            =   100
      TabIndex        =   36
      Top             =   3300
      Width           =   4800
      Begin VB.CommandButton RepAD 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Решить"
         Height          =   1000
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1400
         Width           =   1200
      End
      Begin VB.TextBox pMre 
         Height          =   450
         Left            =   1000
         TabIndex        =   44
         Text            =   "0"
         Top             =   3000
         Width           =   1000
      End
      Begin VB.TextBox pDre 
         Height          =   450
         Left            =   1000
         TabIndex        =   43
         Text            =   "0"
         Top             =   2200
         Width           =   1500
      End
      Begin VB.TextBox pAre 
         Height          =   450
         Left            =   1000
         TabIndex        =   42
         Text            =   "0"
         Top             =   1400
         Width           =   1500
      End
      Begin VB.ComboBox pNKP 
         Height          =   450
         ItemData        =   "REPER.frx":0040
         Left            =   1000
         List            =   "REPER.frx":0053
         TabIndex        =   38
         Text            =   "1"
         Top             =   600
         Width           =   1000
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мц="
         Height          =   400
         Left            =   100
         TabIndex        =   41
         Top             =   3000
         Width           =   600
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Д="
         Height          =   400
         Left            =   100
         TabIndex        =   40
         Top             =   2200
         Width           =   500
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         Height          =   375
         Left            =   100
         TabIndex        =   39
         Top             =   1400
         Width           =   500
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ КП"
         Height          =   400
         Left            =   100
         TabIndex        =   37
         Top             =   600
         Width           =   800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Пристрелянные установки"
      Height          =   3100
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   17500
      Begin VB.CommandButton pernaBat 
         BackColor       =   &H0080FF80&
         Caption         =   "Передать поправки на батарею"
         Height          =   1000
         Left            =   12720
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1600
         Width           =   3000
      End
      Begin VB.ComboBox pNbat 
         Height          =   450
         ItemData        =   "REPER.frx":0066
         Left            =   16200
         List            =   "REPER.frx":0073
         TabIndex        =   72
         Text            =   "1"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pN 
         Height          =   450
         Left            =   14800
         TabIndex        =   35
         Text            =   "0"
         Top             =   500
         Width           =   1000
      End
      Begin VB.TextBox pvdN 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   9900
         TabIndex        =   33
         Text            =   "0"
         Top             =   2300
         Width           =   1000
      End
      Begin VB.TextBox pvdDov 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   8200
         TabIndex        =   31
         Text            =   "0"
         Top             =   2300
         Width           =   1000
      End
      Begin VB.TextBox pvdD 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   6250
         TabIndex        =   29
         Text            =   "0"
         Top             =   2300
         Width           =   1000
      End
      Begin VB.TextBox pvDpris 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   4320
         TabIndex        =   27
         Text            =   "0"
         Top             =   2300
         Width           =   1300
      End
      Begin VB.TextBox pvPrpris 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   1800
         TabIndex        =   25
         Text            =   "0"
         Top             =   2300
         Width           =   1000
      End
      Begin VB.TextBox pvdh 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   10900
         TabIndex        =   23
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvDovrt 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   9100
         TabIndex        =   21
         Text            =   "0"
         Top             =   1600
         Width           =   1200
      End
      Begin VB.TextBox pvDrt 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   6600
         TabIndex        =   19
         Text            =   "0"
         Top             =   1600
         Width           =   1200
      End
      Begin VB.TextBox pvhre 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   4800
         TabIndex        =   17
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvYre 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   2700
         TabIndex        =   15
         Text            =   "0"
         Top             =   1600
         Width           =   1500
      End
      Begin VB.TextBox pvXre 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   600
         TabIndex        =   13
         Text            =   "0"
         Top             =   1600
         Width           =   1500
      End
      Begin VB.TextBox pDov 
         Height          =   450
         Left            =   12300
         TabIndex        =   10
         Text            =   "0"
         Top             =   500
         Width           =   1200
      End
      Begin VB.TextBox ppric 
         Height          =   450
         Left            =   9100
         TabIndex        =   8
         Text            =   "0"
         Top             =   500
         Width           =   1000
      End
      Begin VB.ComboBox pvzriv 
         Height          =   450
         ItemData        =   "REPER.frx":0080
         Left            =   6800
         List            =   "REPER.frx":008D
         TabIndex        =   6
         Text            =   "РГМ"
         Top             =   500
         Width           =   1000
      End
      Begin VB.ComboBox psnar 
         Height          =   450
         ItemData        =   "REPER.frx":00A4
         Left            =   4000
         List            =   "REPER.frx":00AE
         TabIndex        =   4
         Text            =   "ОФ"
         Top             =   500
         Width           =   1000
      End
      Begin VB.ComboBox pzar 
         Height          =   450
         ItemData        =   "REPER.frx":00BA
         Left            =   1200
         List            =   "REPER.frx":00D0
         TabIndex        =   2
         Text            =   "Полн"
         Top             =   500
         Width           =   1500
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Бат"
         Height          =   405
         Left            =   16200
         TabIndex        =   71
         Top             =   1000
         Width           =   1005
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Трубка"
         Height          =   400
         Left            =   13700
         TabIndex        =   34
         Top             =   500
         Width           =   1000
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dN"
         Height          =   400
         Left            =   9300
         TabIndex        =   32
         Top             =   2300
         Width           =   500
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dДов."
         Height          =   400
         Left            =   7350
         TabIndex        =   30
         Top             =   2300
         Width           =   800
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД"
         Height          =   400
         Left            =   5700
         TabIndex        =   28
         Top             =   2300
         Width           =   500
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Д. пристр."
         Height          =   400
         Left            =   2900
         TabIndex        =   26
         Top             =   2300
         Width           =   1400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Уг. возвыш."
         Height          =   400
         Left            =   100
         TabIndex        =   24
         Top             =   2300
         Width           =   1600
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dh="
         Height          =   405
         Left            =   10400
         TabIndex        =   22
         Top             =   1600
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дов.Р.т="
         Height          =   405
         Left            =   7900
         TabIndex        =   20
         Top             =   1600
         Width           =   1200
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дрт="
         Height          =   405
         Left            =   5900
         TabIndex        =   18
         Top             =   1600
         Width           =   705
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   400
         Left            =   4300
         TabIndex        =   16
         Top             =   1600
         Width           =   500
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2200
         TabIndex        =   14
         Top             =   1600
         Width           =   500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Х="
         Height          =   400
         Left            =   100
         TabIndex        =   12
         Top             =   1600
         Width           =   500
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ДОКЛАД"
         Height          =   400
         Left            =   200
         TabIndex        =   11
         Top             =   1100
         Width           =   1300
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Доворот от ОН"
         Height          =   400
         Left            =   10200
         TabIndex        =   9
         Top             =   500
         Width           =   2100
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Прицел"
         Height          =   400
         Left            =   7900
         TabIndex        =   7
         Top             =   500
         Width           =   1100
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Взрыватель"
         Height          =   400
         Left            =   5100
         TabIndex        =   5
         Top             =   500
         Width           =   1700
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Снаряд"
         Height          =   400
         Left            =   2900
         TabIndex        =   3
         Top             =   500
         Width           =   1000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Заряд"
         Height          =   400
         Left            =   240
         TabIndex        =   1
         Top             =   500
         Width           =   855
      End
   End
End
Attribute VB_Name = "REPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
REPER.Hide
End Sub

Private Sub Command3_Click()
Dim snar As String, vzriv As String, zar As String
Dim snvz As Single, Pric As Single, N As Single, dov As Single, Nop As Single
Dim Xb As Single, Yb As Single, hb As Single, OH As Single
snar = psnar: vzriv = pVzriv
If snar = "ОФ" And vzriv = "РГМ" Then snvz = 0
If snar = "ОФ" And vzriv = "АР-5" Then snvz = 1
If snar = "ОФ" And vzriv = "В-90" Then snvz = 2
If snar = "3Ш" And vzriv = "ДТМ-75" Then snvz = 3
zar = pzar: Pric = ppric: N = pN: dov = pDov: Nop = pNbat
If Nop = 1 Then OH = BP.pOH1
If Nop = 2 Then OH = BP.pOH2
If Nop = 3 Then OH = BP.pOH3
Xre = pXre: Yre = pYre: hre = phre
If Nop = 1 Then Xb = BP.pX1: Yb = BP.pY1: hb = BP.ph1
If Nop = 2 Then Xb = BP.pX2: Yb = BP.pY2: hb = BP.ph2
If Nop = 3 Then Xb = BP.pX3: Yb = BP.pY3: hb = BP.ph3
 dxre = Xre - Xb: dyre = Yre - Yb: dhre = hre - hb
 Dret = Sqr(dxre ^ 2 + dyre ^ 2)
 Are = Abs(Atn(dyre / (dxre + 0.1)) / 3.141592 * 30) * 100
 Yrre = dhre / (Dret * 0.001 + 0.01) * 0.95
  If dxre > 0 And dyre > 0 Then Ygolre = Int(Are)
  If dxre < 0 And dyre > 0 Then Ygolre = Int(3000 - Are)
  If dxre < 0 And dyre < 0 Then Ygolre = Int(3000 + Are)
  If dxre > 0 And dyre < 0 Then Ygolre = Int(6000 - Are)
  If OH <= 1500 And Ygolre >= 4500 Then
         dovret = Ygolre - (OH + 6000)
     Else
         dovret = Ygolre - (OH)
       End If
KPEnast zar, Pric, Yrre, kpe
    daep = Yrre * kpe: Pricpris = Pric - Yrre - daep
   If zar = "Полн" Then
   Open App.Path & "\2C1P" For Input As #1
   ElseIf zar = "Умен" Then
   Open App.Path & "\2C1Y" For Input As #1
   ElseIf zar = "Перв" Then
   Open App.Path & "\2C11" For Input As #1
   ElseIf zar = "Втор" Then
   Open App.Path & "\2C12" For Input As #1
   ElseIf zar = "Трет" Then
   Open App.Path & "\2C13" For Input As #1
   ElseIf zar = "Четверт" Then
   Open App.Path & "\2C14" For Input As #1
   Else
   Open App.Path & "\2C1P" For Input As #1
End If
224 If EOF(1) Then GoTo 227
225 Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8, ta9, ta10, ta11, ta12, ta13, ta14, ta15, ta16, ta17, ta18, ta19, ta20, ta21, ta22, ta23, ta24, ta25, ta26, ta27, ta28, ta29, ta30, ta31, ta32, ta33, ta34, ta35, ta36, ta37
226 If ta2 <= Pricpris Then GoTo 22611
 GoTo 224
22611 Prta = ta2: dXtus = ta3: Dta = ta1
227 Close #1
  Dpris = (Pricpris - Prta) * dXtus + Dta
REPERN snvz, zar, Pricpris, dN, N, dXtus, Dprisn
If snvz = 2 Or snvz = 3 Then
        Dpris = Dprisn
        Else
        Dpris = Dpris
End If
dDr = Dpris - Dret: ddovre = dov - dovret
pvXre.Text = Xre: pvYre.Text = Yre: pvhre.Text = hre: pvDrt.Text = Round(Dret): pvDovrt.Text = Round(dovret): pvdh.Text = Round(dhre)
pvPrpris.Text = Round(Pricpris): pvDpris.Text = Round(Dpris): pvdD.Text = Round(dDr): pvdDov.Text = Round(ddovre): pvdN.Text = Round(dN, 1)
End Sub
Function KPEnast(ByVal zar As String, ByVal Pric As Single, ByVal Yrre As Single, kpe) As Single
223  If zar = "Полн" Then
        Open App.Path & "\2c1kpep" For Input As #1
            ElseIf zar = "Умен" Then
            Open App.Path & "\2c1kpey" For Input As #1
                ElseIf zar = "Перв" Then
                Open App.Path & "\2c1kpe1" For Input As #1
                    ElseIf zar = "Втор" Then
                    Open App.Path & "\2c1kpe2" For Input As #1
                        ElseIf zar = "Трет" Then
                        Open App.Path & "\2c1kpe3" For Input As #1
                            ElseIf zar = "Четверт" Then
                            Open App.Path & "\2c1kpe4" For Input As #1
                                Else
                                Open App.Path & "\2c1kpep" For Input As #1
        End If
22311 If EOF(1) Then GoTo 22312
  Input #1, e1, e2, e3, e4, e5
   If Pric >= e1 And e1 + 20 >= Pric Then kpew = e2: kpen = e3: GoTo 22312
   GoTo 22311
22312 Close #1
    If Yrr >= 0 Then kpe = kpew * 0.1
    If Yrr < 0 Then kpe = kpen * 0.1
End Function
Function REPERN(ByVal snvz As Single, ByVal zar As String, ByVal Pricpris As Single, dN, ByVal N As Single, ByVal dXtus As Single, Dprisn) As Single
If snvz = 2 Then
       If zar = "Полн" Then
       Open App.Path & "\B-90P" For Input As #1
       ElseIf zar = "Умен" Then
       Open App.Path & "\B-90y" For Input As #1
       ElseIf zar = "Перв" Then
       Open App.Path & "\B-901" For Input As #1
       ElseIf zar = "Втор" Then
       Open App.Path & "\B-902" For Input As #1
       ElseIf zar = "Трет" Then
       Open App.Path & "\B-903" For Input As #1
       Else
       Open App.Path & "\B-90P" For Input As #1
       End If
       ElseIf snvz = 0 Then
       GoTo 2261001
       Else
       If zar = 0 Then Open App.Path & "\3sh-P" For Input As #1
       If zar = 5 Then Open App.Path & "\3SH-Y" For Input As #1
End If
22511 If EOF(1) Then GoTo 22515
Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8
If ta2 >= Pricpris Then GoTo 22512
GoTo 22511
22512   Dta = ta1: Prta = ta2: dNtus = ta4: Nta = ta3: GoTo 22515
22515 Close
Nrp = (Pricpris - Prta) * dNtus + Nta
dN = N - Nrp - (daep * dNtus)
Dprisn = (Pricpris - Prta) * dXtus + Dta
2261001
End Function

Private Sub pernaBat_Click()
Dim p As String
Dim Nbat As Single
Nbat = pNbat
If Nbat = 1 Then
    pvZar1.Text = pzar: pvsn1.Text = psnar: pvvzr1.Text = pVzriv: pvDr1.Text = pvDrt: pvdD1.Text = pvdD: pvdDov1.Text = pvdDov: pvdN1.Text = pvdN
    ElseIf Nbat = 2 Then
        pvZar2.Text = pzar: pvsn2.Text = psnar: pvvzr2.Text = pVzriv: pvDr2.Text = pvDrt: pvdD2.Text = pvdD: pvdDov2.Text = pvdDov: pvdN2.Text = pvdN
        Else
            pvZar3.Text = pzar: pvsn3.Text = psnar: pvvzr3.Text = pVzriv: pvDr3.Text = pvDrt: pvdD3.Text = pvdD: pvdDov3.Text = pvdDov: pvdN3.Text = pvdN
End If
p = Now
If Nbat = 1 Then
    ptime1.Text = p
    ElseIf Nbat = 2 Then
        ptime2.Text = p
        Else
            ptime3.Text = p
End If
End Sub

Private Sub ppric_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pDov.Text = ""
pDov.SetFocus
End If
End Sub
Private Sub pDov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pN.Text = ""
pN.SetFocus
End If
End Sub
Private Sub pAre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pDre.Text = ""
pDre.SetFocus
End If
End Sub
Private Sub pDre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMre.Text = ""
pMre.SetFocus
End If
End Sub
Private Sub pAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMl.Text = ""
pMl.SetFocus
End If
End Sub
Private Sub pAp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMp.Text = ""
pMp.SetFocus
End If
End Sub
Private Sub pXre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYre.Text = ""
pYre.SetFocus
End If
End Sub
Private Sub pYre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phre.Text = ""
phre.SetFocus
End If
End Sub

Private Sub RepAA_Click()
Dim snar As String, vzriv As String, zar As String
Dim snvz As Single, Pric As Single, N As Single, dov As Single, Nop As Single
Dim Xb As Single, Yb As Single, hb As Single, OH As Single
Dim Alev As Single, Mcl As Single, nnpl As Single, Aprav As Single, Mcp As Single, nnpp As Single
Dim Xl As Single, Yl As Single, hl As Single, Xp As Single, Yp As Single, hp As Single
snar = psnar: vzriv = pVzriv
If snar = "ОФ" And vzriv = "РГМ" Then snvz = 0
If snar = "ОФ" And vzriv = "АР-5" Then snvz = 1
If snar = "ОФ" And vzriv = "В-90" Then snvz = 2
If snar = "3Ш" And vzriv = "ДТМ-75" Then snvz = 3
zar = pzar: Pric = ppric: N = pN: dov = pDov: Nop = pNbat
If Nop = 1 Then OH = BP.pOH1
If Nop = 2 Then OH = BP.pOH2
If Nop = 3 Then OH = BP.pOH3
    Alev = pAl: Mcl = pMl: nnpl = pNKPL
  Aprav = pAp: Mcp = pMp: nnpp = pNKPP
'LEV
  If nnpl = 1 Then Xl = BP.pXkp1: Yl = BP.pYkp1: hl = BP.phkp1
  If nnpl = 2 Then Xl = BP.pXkp2: Yl = BP.pYkp2: hl = BP.phkp2
  If nnpl = 3 Then Xl = BP.pXkp3: Yl = BP.pYkp3: hl = BP.phkp3
  If nnpl = 4 Then Xl = BP.pXkp4: Yl = BP.pYkp4: hl = BP.phkp4
    If nnpl = 5 Then Xl = BP.pXkp5: Yl = BP.pYkp5: hl = BP.phkp5
'PRAV
  If nnpp = 1 Then Xp = BP.pXkp1: Yp = BP.pYkp1: hp = BP.phkp1
  If nnpp = 2 Then Xp = BP.pXkp2: Yp = BP.pYkp2: hp = BP.phkp2
  If nnpp = 3 Then Xp = BP.pXkp3: Yp = BP.pYkp3: hp = BP.phkp3
  If nnpp = 4 Then Xp = BP.pXkp4: Yp = BP.pYkp4: hp = BP.phkp4
  If nnpp = 5 Then Xp = BP.pXkp5: Yp = BP.pYkp5: hp = BP.phkp5
dxso = Xp - Xl: dyso = Yp - Yl
  baz = Sqr(dxso ^ 2 + dyso ^ 2)
  aso = Abs(Atn(dyso / (dxso + 0.1)) / 3.141592 * 30) * 100

  If dxso > 0 And dyso > 0 Then Ygolbaz = Int(aso)
  If dxso < 0 And dyso > 0 Then Ygolbaz = Int(3000 - aso)
  If dxso < 0 And dyso < 0 Then Ygolbaz = Int(3000 + aso)
  If dxso > 0 And dyso < 0 Then Ygolbaz = Int(6000 - aso)
  If Alev < 1500 And Aprav > 4500 Then
  fi = Abs(Alev + 6000 - Aprav)
  ElseIf Alev > 4500 And Aprav < 1500 Then
  fi = Abs(Alev - (Aprav + 6000))
  Else
   fi = Abs(Alev - Aprav)
  End If
  If Alev < 1500 And Ygolbaz > 4500 Then
  blev = Abs(Alev + 6000 - Ygolbaz)
  ElseIf Alev > 4500 And Ygolbaz < 1500 Then
  blev = Abs(Alev - (Ygolbaz + 6000))
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
 
  Dlev = baz / (Sin(fi / 100 * 6 * 3.141592 / 180) + 0.001) * Sin(bprav / 100 * 6 * 3.141592 / 180)
  Dprav = baz / (Sin(fi / 100 * 6 * 3.141592 / 180) + 0.001) * Sin(blev / 100 * 6 * 3.141592 / 180)
  Xcso = Cos(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Xl
  Ycso = Sin(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Yl
  If Mcl = 0 Then hc = Mcp * (Dprav * 0.001) * 1.05 + hp
  If Mcp = 0 Then hc = Mcl * (Dlev * 0.001) * 1.05 + hl
  xc = Xcso: yc = Ycso
    Xre = xc: Yre = yc: hre = hc
If Nop = 1 Then Xb = BP.pX1: Yb = BP.pY1: hb = BP.ph1
If Nop = 2 Then Xb = BP.pX2: Yb = BP.pY2: hb = BP.ph2
If Nop = 3 Then Xb = BP.pX3: Yb = BP.pY3: hb = BP.ph3
 dxre = Xre - Xb: dyre = Yre - Yb: dhre = hre - hb
 Dret = Sqr(dxre ^ 2 + dyre ^ 2)
 Are = Abs(Atn(dyre / (dxre + 0.1)) / 3.141592 * 30) * 100
 Yrre = dhre / (Dret * 0.001 + 0.01) * 0.95
  If dxre > 0 And dyre > 0 Then Ygolre = Int(Are)
  If dxre < 0 And dyre > 0 Then Ygolre = Int(3000 - Are)
  If dxre < 0 And dyre < 0 Then Ygolre = Int(3000 + Are)
  If dxre > 0 And dyre < 0 Then Ygolre = Int(6000 - Are)
  If OH <= 1500 And Ygolre >= 4500 Then
         dovret = Ygolre - (OH + 6000)
     Else
         dovret = Ygolre - (OH)
       End If
KPEnast zar, Pric, Yrre, kpe
    daep = Yrre * kpe: Pricpris = Pric - Yrre - daep
   If zar = "Полн" Then
   Open App.Path & "\2C1P" For Input As #1
   ElseIf zar = "Умен" Then
   Open App.Path & "\2C1Y" For Input As #1
   ElseIf zar = "Перв" Then
   Open App.Path & "\2C11" For Input As #1
   ElseIf zar = "Втор" Then
   Open App.Path & "\2C12" For Input As #1
   ElseIf zar = "Трет" Then
   Open App.Path & "\2C13" For Input As #1
   ElseIf zar = "Четверт" Then
   Open App.Path & "\2C14" For Input As #1
   Else
   Open App.Path & "\2C1P" For Input As #1
End If
224 If EOF(1) Then GoTo 227
225 Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8, ta9, ta10, ta11, ta12, ta13, ta14, ta15, ta16, ta17, ta18, ta19, ta20, ta21, ta22, ta23, ta24, ta25, ta26, ta27, ta28, ta29, ta30, ta31, ta32, ta33, ta34, ta35, ta36, ta37
226 If ta2 <= Pricpris Then GoTo 22611
 GoTo 224
22611 Prta = ta2: dXtus = ta3: Dta = ta1
227 Close #1
  Dpris = (Pricpris - Prta) * dXtus + Dta
REPERN snvz, zar, Pricpris, dN, N, dXtus, Dprisn
If snvz = 2 Or snvz = 3 Then
        Dpris = Dprisn
        Else
        Dpris = Dpris
End If
dDr = Dpris - Dret: ddovre = dov - dovret
pvXre.Text = Round(Xre): pvYre.Text = Round(Yre): pvhre.Text = Round(hre): pvDrt.Text = Round(Dret): pvDovrt.Text = Round(dovret): pvdh.Text = Round(dhre)
pvPrpris.Text = Round(Pricpris): pvDpris.Text = Round(Dpris): pvdD.Text = Round(dDr): pvdDov.Text = Round(ddovre): pvdN.Text = Round(dN, 1)

End Sub

Private Sub RepAD_Click()
Dim snar As String, vzriv As String, zar As String
Dim snvz As Single, Pric As Single, N As Single, dov As Single, Nop As Single, nkp As Single
Dim Are As Single, Dre As Single, Mcre As Single, xkp As Single, ykp As Single, hkp As Single
Dim Xb As Single, Yb As Single, hb As Single, OH As Single
snar = psnar: vzriv = pVzriv
If snar = "ОФ" And vzriv = "РГМ" Then snvz = 0
If snar = "ОФ" And vzriv = "АР-5" Then snvz = 1
If snar = "ОФ" And vzriv = "В-90" Then snvz = 2
If snar = "3Ш" And vzriv = "ДТМ-75" Then snvz = 3
zar = pzar: Pric = ppric: N = pN: dov = pDov: Nop = pNbat
If Nop = 1 Then OH = BP.pOH1
If Nop = 2 Then OH = BP.pOH2
If Nop = 3 Then OH = BP.pOH3
Are = pAre: Dre = pDre: Mcre = pMre: nkp = pNKP
   If nkp = 1 Then xkp = BP.pXkp1: ykp = BP.pYkp1: hkp = BP.phkp1
   If nkp = 2 Then xkp = BP.pXkp2: ykp = BP.pYkp2: hkp = BP.phkp2
   If nkp = 3 Then xkp = BP.pXkp3: ykp = BP.pYkp3: hkp = BP.phkp3
   If nkp = 4 Then xkp = BP.pXkp4: ykp = BP.pYkp4: hkp = BP.phkp4
      If nkp = 5 Then xkp = BP.pXkp5: ykp = BP.pYkp5: hkp = BP.phkp5
 Xre = Cos(Are / 100 * 6 * 3.141592 / 180) * Dre + xkp
 Yre = Sin(Are / 100 * 6 * 3.141592 / 180) * Dre + ykp
 hre = Mcre * (Dre * 0.001) * 1.05 + hkp
If Nop = 1 Then Xb = BP.pX1: Yb = BP.pY1: hb = BP.ph1
If Nop = 2 Then Xb = BP.pX2: Yb = BP.pY2: hb = BP.ph2
If Nop = 3 Then Xb = BP.pX3: Yb = BP.pY3: hb = BP.ph3
 dxre = Xre - Xb: dyre = Yre - Yb: dhre = hre - hb
 Dret = Sqr(dxre ^ 2 + dyre ^ 2)
 Are = Abs(Atn(dyre / (dxre + 0.1)) / 3.141592 * 30) * 100
 Yrre = dhre / (Dret * 0.001 + 0.01) * 0.95
  If dxre > 0 And dyre > 0 Then Ygolre = Int(Are)
  If dxre < 0 And dyre > 0 Then Ygolre = Int(3000 - Are)
  If dxre < 0 And dyre < 0 Then Ygolre = Int(3000 + Are)
  If dxre > 0 And dyre < 0 Then Ygolre = Int(6000 - Are)
  If OH <= 1500 And Ygolre >= 4500 Then
         dovret = Ygolre - (OH + 6000)
     Else
         dovret = Ygolre - (OH)
       End If
KPEnast zar, Pric, Yrre, kpe
    daep = Yrre * kpe: Pricpris = Pric - Yrre - daep
   If zar = "Полн" Then
   Open App.Path & "\2C1P" For Input As #1
   ElseIf zar = "Умен" Then
   Open App.Path & "\2C1Y" For Input As #1
   ElseIf zar = "Перв" Then
   Open App.Path & "\2C11" For Input As #1
   ElseIf zar = "Втор" Then
   Open App.Path & "\2C12" For Input As #1
   ElseIf zar = "Трет" Then
   Open App.Path & "\2C13" For Input As #1
   ElseIf zar = "Четверт" Then
   Open App.Path & "\2C14" For Input As #1
   Else
   Open App.Path & "\2C1P" For Input As #1
End If
224 If EOF(1) Then GoTo 227
225 Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8, ta9, ta10, ta11, ta12, ta13, ta14, ta15, ta16, ta17, ta18, ta19, ta20, ta21, ta22, ta23, ta24, ta25, ta26, ta27, ta28, ta29, ta30, ta31, ta32, ta33, ta34, ta35, ta36, ta37
226 If ta2 <= Pricpris Then GoTo 22611
 GoTo 224
22611 Prta = ta2: dXtus = ta3: Dta = ta1
227 Close #1
  Dpris = (Pricpris - Prta) * dXtus + Dta
REPERN snvz, zar, Pricpris, dN, N, dXtus, Dprisn
If snvz = 2 Or snvz = 3 Then
        Dpris = Dprisn
        Else
        Dpris = Dpris
End If
dDr = Dpris - Dret: ddovre = dov - dovret
pvXre.Text = Round(Xre): pvYre.Text = Round(Yre): pvhre.Text = Round(hre): pvDrt.Text = Round(Dret): pvDovrt.Text = Round(dovret): pvdh.Text = Round(dhre)
pvPrpris.Text = Round(Pricpris): pvDpris.Text = Round(Dpris): pvdD.Text = Round(dDr): pvdDov.Text = Round(ddovre): pvdN.Text = Round(dN, 1)

End Sub
