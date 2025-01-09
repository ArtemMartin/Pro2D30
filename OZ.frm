VERSION 5.00
Begin VB.Form OZ 
   Caption         =   "Огневая задача"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
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
   Moveable        =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   Begin VB.CommandButton okPopravki 
      BackColor       =   &H00FF8080&
      Caption         =   "Команда"
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
      Height          =   1100
      Left            =   18100
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   8600
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Справка"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   16300
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   8600
      Width           =   1300
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Подготовка стрельбы"
      Height          =   2100
      Left            =   14000
      TabIndex        =   153
      Top             =   5950
      Width           =   6400
      Begin VB.ComboBox pRep3 
         Height          =   405
         ItemData        =   "OZ.frx":0000
         Left            =   4800
         List            =   "OZ.frx":000A
         TabIndex        =   160
         Text            =   "Полная"
         Top             =   1100
         Width           =   1500
      End
      Begin VB.ComboBox pRep2 
         Height          =   405
         ItemData        =   "OZ.frx":0022
         Left            =   3200
         List            =   "OZ.frx":002C
         TabIndex        =   159
         Text            =   "Полная"
         Top             =   1100
         Width           =   1500
      End
      Begin VB.ComboBox pRep1 
         Height          =   405
         ItemData        =   "OZ.frx":0044
         Left            =   1600
         List            =   "OZ.frx":004E
         TabIndex        =   158
         Text            =   "Полная"
         Top             =   1100
         Width           =   1500
      End
      Begin VB.Label Label53 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 Бат"
         Height          =   300
         Left            =   5100
         TabIndex        =   157
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label52 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 Бат"
         Height          =   300
         Left            =   3500
         TabIndex        =   156
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label51 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 Бат"
         Height          =   300
         Left            =   1900
         TabIndex        =   155
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label50 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Подготовка"
         Height          =   300
         Left            =   240
         TabIndex        =   154
         Top             =   1100
         Width           =   1400
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Выбор вид стрельбы"
      Height          =   2100
      Left            =   14000
      TabIndex        =   145
      Top             =   3700
      Width           =   6400
      Begin VB.ComboBox pStre3 
         Height          =   405
         ItemData        =   "OZ.frx":0066
         Left            =   4800
         List            =   "OZ.frx":0070
         TabIndex        =   152
         Text            =   "Настильная"
         Top             =   1100
         Width           =   1500
      End
      Begin VB.ComboBox pStre2 
         Height          =   405
         ItemData        =   "OZ.frx":008B
         Left            =   3200
         List            =   "OZ.frx":0095
         TabIndex        =   151
         Text            =   "Настильная"
         Top             =   1100
         Width           =   1500
      End
      Begin VB.ComboBox pStre1 
         Height          =   405
         ItemData        =   "OZ.frx":00B0
         Left            =   1600
         List            =   "OZ.frx":00BA
         TabIndex        =   150
         Text            =   "Настильная"
         Top             =   1100
         Width           =   1500
      End
      Begin VB.Label Label49 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Стрельба"
         Height          =   300
         Left            =   240
         TabIndex        =   149
         Top             =   1100
         Width           =   1200
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 Бат"
         Height          =   300
         Left            =   5100
         TabIndex        =   148
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label47 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 Бат"
         Height          =   300
         Left            =   3500
         TabIndex        =   147
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label46 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 Бат"
         Height          =   300
         Left            =   1900
         TabIndex        =   146
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Выбор Снаряд, Взрыватель, Заряд"
      Height          =   3495
      Left            =   14000
      TabIndex        =   129
      Top             =   100
      Width           =   6400
      Begin VB.ComboBox pSnar1 
         Height          =   405
         ItemData        =   "OZ.frx":00D5
         Left            =   1600
         List            =   "OZ.frx":00EB
         TabIndex        =   138
         Text            =   "ОФ"
         Top             =   1100
         Width           =   1300
      End
      Begin VB.ComboBox pVzr1 
         Height          =   405
         ItemData        =   "OZ.frx":010C
         Left            =   1600
         List            =   "OZ.frx":011F
         TabIndex        =   137
         Text            =   "РГМ"
         Top             =   1800
         Width           =   1300
      End
      Begin VB.ComboBox pZar1 
         Height          =   405
         ItemData        =   "OZ.frx":0142
         Left            =   1600
         List            =   "OZ.frx":0158
         TabIndex        =   136
         Text            =   "Полн"
         Top             =   2500
         Width           =   1300
      End
      Begin VB.ComboBox pSnar2 
         Height          =   405
         ItemData        =   "OZ.frx":0183
         Left            =   3200
         List            =   "OZ.frx":0199
         TabIndex        =   135
         Text            =   "ОФ"
         Top             =   1100
         Width           =   1300
      End
      Begin VB.ComboBox pVzr2 
         Height          =   405
         ItemData        =   "OZ.frx":01BA
         Left            =   3200
         List            =   "OZ.frx":01CD
         TabIndex        =   134
         Text            =   "РГМ"
         Top             =   1800
         Width           =   1300
      End
      Begin VB.ComboBox pZar2 
         Height          =   405
         ItemData        =   "OZ.frx":01F0
         Left            =   3200
         List            =   "OZ.frx":0206
         TabIndex        =   133
         Text            =   "Полн"
         Top             =   2500
         Width           =   1300
      End
      Begin VB.ComboBox pSnar3 
         Height          =   405
         ItemData        =   "OZ.frx":0231
         Left            =   4800
         List            =   "OZ.frx":0247
         TabIndex        =   132
         Text            =   "ОФ"
         Top             =   1100
         Width           =   1300
      End
      Begin VB.ComboBox pVzr3 
         Height          =   405
         ItemData        =   "OZ.frx":0268
         Left            =   4800
         List            =   "OZ.frx":027B
         TabIndex        =   131
         Text            =   "РГМ"
         Top             =   1800
         Width           =   1300
      End
      Begin VB.ComboBox pZar3 
         Height          =   405
         ItemData        =   "OZ.frx":029E
         Left            =   4800
         List            =   "OZ.frx":02B4
         TabIndex        =   130
         Text            =   "Полн"
         Top             =   2500
         Width           =   1300
      End
      Begin VB.Label Label62 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 Бат"
         Height          =   300
         Left            =   1900
         TabIndex        =   144
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label63 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 Бат"
         Height          =   300
         Left            =   3500
         TabIndex        =   143
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label64 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 Бат"
         Height          =   300
         Left            =   5100
         TabIndex        =   142
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label65 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Снаряд"
         Height          =   300
         Left            =   100
         TabIndex        =   141
         Top             =   1100
         Width           =   1500
      End
      Begin VB.Label Label66 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Взрыватель"
         Height          =   300
         Left            =   100
         TabIndex        =   140
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label67 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Заряд"
         Height          =   300
         Left            =   100
         TabIndex        =   139
         Top             =   2500
         Width           =   1500
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
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
      Height          =   1100
      Left            =   14500
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   8600
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Данные по цели"
      Height          =   9615
      Left            =   7400
      TabIndex        =   40
      Top             =   100
      Width           =   6495
      Begin VB.TextBox pvdDov3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   4400
         TabIndex        =   127
         Text            =   "0"
         Top             =   8700
         Width           =   1200
      End
      Begin VB.TextBox pvDisch3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   4400
         TabIndex        =   126
         Text            =   "0"
         Top             =   8300
         Width           =   1200
      End
      Begin VB.TextBox pvdDov2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   3100
         TabIndex        =   125
         Text            =   "0"
         Top             =   8700
         Width           =   1200
      End
      Begin VB.TextBox pvDisch2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   3100
         TabIndex        =   124
         Text            =   "0"
         Top             =   8300
         Width           =   1200
      End
      Begin VB.TextBox pvdDov1 
         BackColor       =   &H00C0C0FF&
         Height          =   405
         Left            =   1800
         TabIndex        =   123
         Text            =   "0"
         Top             =   8700
         Width           =   1200
      End
      Begin VB.TextBox pvdD3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   4400
         TabIndex        =   122
         Text            =   "0"
         Top             =   7900
         Width           =   1200
      End
      Begin VB.TextBox pvOH3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   4400
         TabIndex        =   121
         Text            =   "0"
         Top             =   7500
         Width           =   1200
      End
      Begin VB.TextBox pvYr3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   4400
         TabIndex        =   120
         Text            =   "0"
         Top             =   7100
         Width           =   1200
      End
      Begin VB.TextBox pvdD2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   3100
         TabIndex        =   119
         Text            =   "0"
         Top             =   7900
         Width           =   1200
      End
      Begin VB.TextBox pvOH2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   3100
         TabIndex        =   118
         Text            =   "0"
         Top             =   7500
         Width           =   1200
      End
      Begin VB.TextBox pvYr2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   3100
         TabIndex        =   117
         Text            =   "0"
         Top             =   7100
         Width           =   1200
      End
      Begin VB.TextBox pvDisch1 
         BackColor       =   &H00C0C0FF&
         Height          =   405
         Left            =   1800
         TabIndex        =   116
         Text            =   "0"
         Top             =   8300
         Width           =   1200
      End
      Begin VB.TextBox pvdD1 
         BackColor       =   &H00C0C0FF&
         Height          =   405
         Left            =   1800
         TabIndex        =   115
         Text            =   "0"
         Top             =   7900
         Width           =   1200
      End
      Begin VB.TextBox pvOH1 
         BackColor       =   &H00C0C0FF&
         Height          =   405
         Left            =   1800
         TabIndex        =   114
         Text            =   "0"
         Top             =   7500
         Width           =   1200
      End
      Begin VB.TextBox pvDovt3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   113
         Text            =   "0"
         Top             =   6700
         Width           =   1200
      End
      Begin VB.TextBox pvYgt3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   112
         Text            =   "0"
         Top             =   6300
         Width           =   1200
      End
      Begin VB.TextBox pvDt3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   111
         Text            =   "0"
         Top             =   5900
         Width           =   1200
      End
      Begin VB.TextBox pvVd3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   110
         Text            =   "0"
         Top             =   5500
         Width           =   1200
      End
      Begin VB.TextBox pvVustra3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   109
         Text            =   "0"
         Top             =   5100
         Width           =   1200
      End
      Begin VB.TextBox pvDovt2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   108
         Text            =   "0"
         Top             =   6700
         Width           =   1200
      End
      Begin VB.TextBox pvYgt2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   107
         Text            =   "0"
         Top             =   6300
         Width           =   1200
      End
      Begin VB.TextBox pvDt2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   106
         Text            =   "0"
         Top             =   5900
         Width           =   1200
      End
      Begin VB.TextBox pvVd2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   105
         Text            =   "0"
         Top             =   5500
         Width           =   1200
      End
      Begin VB.TextBox pvVustra2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   104
         Text            =   "0"
         Top             =   5100
         Width           =   1200
      End
      Begin VB.TextBox pvPolet2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   103
         Text            =   "0"
         Top             =   4700
         Width           =   1200
      End
      Begin VB.TextBox pvYr1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   102
         Text            =   "0"
         Top             =   7100
         Width           =   1200
      End
      Begin VB.TextBox pvDovt1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   101
         Text            =   "0"
         Top             =   6700
         Width           =   1200
      End
      Begin VB.TextBox pvYgt1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   100
         Text            =   "0"
         Top             =   6300
         Width           =   1200
      End
      Begin VB.TextBox pvDt1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   99
         Text            =   "0"
         Top             =   5900
         Width           =   1200
      End
      Begin VB.TextBox pvVd1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   98
         Text            =   "0"
         Top             =   5500
         Width           =   1200
      End
      Begin VB.TextBox pvVustra1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   97
         Text            =   "0"
         Top             =   5100
         Width           =   1200
      End
      Begin VB.TextBox pvPolet3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   86
         Text            =   "0"
         Top             =   4700
         Width           =   1200
      End
      Begin VB.TextBox pvdNtus3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   85
         Text            =   "0"
         Top             =   4300
         Width           =   1200
      End
      Begin VB.TextBox pvdXtus3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   84
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox pvSk3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   83
         Text            =   "0"
         Top             =   3500
         Width           =   1200
      End
      Begin VB.TextBox pvVeer3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   82
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox pvDov3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   81
         Text            =   "0"
         Top             =   2700
         Width           =   1200
      End
      Begin VB.TextBox pvN3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   80
         Text            =   "0"
         Top             =   2300
         Width           =   1200
      End
      Begin VB.TextBox pvPric3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   79
         Text            =   "0"
         Top             =   1900
         Width           =   1200
      End
      Begin VB.TextBox pvZar3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   78
         Text            =   "0"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.TextBox pvVzr3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   77
         Text            =   "0"
         Top             =   1100
         Width           =   1200
      End
      Begin VB.TextBox pvSnar3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   4400
         TabIndex        =   76
         Text            =   "0"
         Top             =   700
         Width           =   1200
      End
      Begin VB.TextBox pvdNtus2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   75
         Text            =   "0"
         Top             =   4300
         Width           =   1200
      End
      Begin VB.TextBox pvdXtus2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   74
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox pvSk2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   73
         Text            =   "0"
         Top             =   3500
         Width           =   1200
      End
      Begin VB.TextBox pvVeer2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   72
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox pvDov2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   71
         Text            =   "0"
         Top             =   2700
         Width           =   1200
      End
      Begin VB.TextBox pvN2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   70
         Text            =   "0"
         Top             =   2300
         Width           =   1200
      End
      Begin VB.TextBox pvPric2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   69
         Text            =   "0"
         Top             =   1900
         Width           =   1200
      End
      Begin VB.TextBox pvZar2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   68
         Text            =   "0"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.TextBox pvVzr2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   67
         Text            =   "0"
         Top             =   1100
         Width           =   1200
      End
      Begin VB.TextBox pvSnar2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   3100
         TabIndex        =   66
         Text            =   "0"
         Top             =   700
         Width           =   1200
      End
      Begin VB.TextBox pvPolet1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   65
         Text            =   "0"
         Top             =   4700
         Width           =   1200
      End
      Begin VB.TextBox pvdNtus1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   64
         Text            =   "0"
         Top             =   4300
         Width           =   1200
      End
      Begin VB.TextBox pvdXtus1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   63
         Text            =   "0"
         Top             =   3900
         Width           =   1200
      End
      Begin VB.TextBox pvSk1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   62
         Text            =   "0"
         Top             =   3500
         Width           =   1200
      End
      Begin VB.TextBox pvVeer1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   61
         Text            =   "0"
         Top             =   3100
         Width           =   1200
      End
      Begin VB.TextBox pvDov1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1800
         TabIndex        =   60
         Text            =   "0"
         Top             =   2700
         Width           =   1200
      End
      Begin VB.TextBox pvN1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1800
         TabIndex        =   59
         Text            =   "0"
         Top             =   2300
         Width           =   1200
      End
      Begin VB.TextBox pvPric1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1800
         TabIndex        =   58
         Text            =   "0"
         Top             =   1920
         Width           =   1200
      End
      Begin VB.TextBox pvZar1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1800
         TabIndex        =   57
         Text            =   "0"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.TextBox pvVzr1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   56
         Text            =   "0"
         Top             =   1100
         Width           =   1200
      End
      Begin VB.TextBox pvSnar1 
         BackColor       =   &H00C0C0FF&
         Height          =   400
         Left            =   1800
         TabIndex        =   55
         Text            =   "0"
         Top             =   700
         Width           =   1200
      End
      Begin VB.Label Label45 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dДов"
         Height          =   300
         Left            =   100
         TabIndex        =   96
         Top             =   8700
         Width           =   1500
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дисч"
         Height          =   300
         Left            =   100
         TabIndex        =   95
         Top             =   8300
         Width           =   1500
      End
      Begin VB.Label Label43 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД"
         Height          =   300
         Left            =   100
         TabIndex        =   94
         Top             =   7900
         Width           =   1500
      End
      Begin VB.Label Label42 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ОН"
         Height          =   300
         Left            =   100
         TabIndex        =   93
         Top             =   7500
         Width           =   1500
      End
      Begin VB.Label Label41 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ур"
         Height          =   300
         Left            =   100
         TabIndex        =   92
         Top             =   7100
         Width           =   1500
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Доворот т."
         Height          =   300
         Left            =   100
         TabIndex        =   91
         Top             =   6700
         Width           =   1500
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Угол т."
         Height          =   300
         Left            =   100
         TabIndex        =   90
         Top             =   6300
         Width           =   1500
      End
      Begin VB.Label Label38 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дт"
         Height          =   300
         Left            =   100
         TabIndex        =   89
         Top             =   5900
         Width           =   1500
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Вд"
         Height          =   300
         Left            =   100
         TabIndex        =   88
         Top             =   5500
         Width           =   1500
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Выс. траект."
         Height          =   300
         Left            =   100
         TabIndex        =   87
         Top             =   5100
         Width           =   1500
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Полетное"
         Height          =   300
         Left            =   120
         TabIndex        =   54
         Top             =   4700
         Width           =   1500
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dNтыс"
         Height          =   300
         Left            =   100
         TabIndex        =   53
         Top             =   4300
         Width           =   1500
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dХтыс"
         Height          =   300
         Left            =   100
         TabIndex        =   52
         Top             =   3900
         Width           =   1500
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Скачек"
         Height          =   300
         Left            =   100
         TabIndex        =   51
         Top             =   3500
         Width           =   1500
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Веер"
         Height          =   300
         Left            =   100
         TabIndex        =   50
         Top             =   3100
         Width           =   1500
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Доворот ОН"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   49
         Top             =   2700
         Width           =   1700
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Трубка"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   48
         Top             =   2300
         Width           =   1500
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Прицел"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   47
         Top             =   1900
         Width           =   1500
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Заряд"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   100
         TabIndex        =   46
         Top             =   1500
         Width           =   1500
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Взрыватель"
         Height          =   300
         Left            =   100
         TabIndex        =   45
         Top             =   1100
         Width           =   1500
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Снаряд"
         Height          =   300
         Left            =   100
         TabIndex        =   44
         Top             =   700
         Width           =   1500
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   3 Бат"
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   4500
         TabIndex        =   43
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   2 Бат"
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   3200
         TabIndex        =   42
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   1 Бат"
         Height          =   300
         Left            =   1900
         TabIndex        =   41
         Top             =   360
         Width           =   1000
      End
   End
   Begin VB.Frame ZasGrZeli 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Цель"
      Height          =   9615
      Left            =   50
      TabIndex        =   0
      Top             =   100
      Width           =   7215
      Begin VB.ComboBox pplZel 
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
         Left            =   4300
         TabIndex        =   167
         Text            =   "0"
         Top             =   2160
         Width           =   2200
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Цель каждому"
         Height          =   1200
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   8000
         Width           =   1500
      End
      Begin VB.CommandButton ZasGrZ 
         BackColor       =   &H0080FF80&
         Caption         =   "Засечка груповой Цели"
         Height          =   1000
         Left            =   4600
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   4200
         Width           =   1800
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   " КНП по Цели"
         Height          =   900
         Left            =   4600
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   3100
         Width           =   1800
      End
      Begin VB.TextBox pGlybinac 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5500
         TabIndex        =   39
         Text            =   "0"
         Top             =   1400
         Width           =   1000
      End
      Begin VB.TextBox pFrontc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5500
         TabIndex        =   38
         Text            =   "0"
         Top             =   900
         Width           =   1000
      End
      Begin VB.CommandButton OZSopr 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Решить"
         Height          =   900
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6300
         Width           =   1100
      End
      Begin VB.CommandButton OZAD 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Решить"
         Height          =   900
         Left            =   2900
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3100
         Width           =   1100
      End
      Begin VB.CommandButton OZXY 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Решить"
         Height          =   900
         Left            =   2500
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   900
         Width           =   1100
      End
      Begin VB.TextBox pMpc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3900
         TabIndex        =   31
         Text            =   "0"
         Top             =   7300
         Width           =   1000
      End
      Begin VB.TextBox pApc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3900
         TabIndex        =   30
         Text            =   "0"
         Top             =   6800
         Width           =   1500
      End
      Begin VB.ComboBox pnkpP 
         Height          =   405
         ItemData        =   "OZ.frx":02DF
         Left            =   3900
         List            =   "OZ.frx":02F2
         TabIndex        =   29
         Text            =   "1"
         Top             =   6300
         Width           =   800
      End
      Begin VB.TextBox pMlc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1100
         TabIndex        =   25
         Text            =   "0"
         Top             =   7300
         Width           =   1000
      End
      Begin VB.TextBox pAlc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1100
         TabIndex        =   24
         Text            =   "0"
         Top             =   6800
         Width           =   1500
      End
      Begin VB.ComboBox pnkpL 
         Height          =   405
         ItemData        =   "OZ.frx":0305
         Left            =   1100
         List            =   "OZ.frx":0318
         TabIndex        =   23
         Text            =   "1"
         Top             =   6300
         Width           =   800
      End
      Begin VB.TextBox pMc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1100
         TabIndex        =   16
         Text            =   "0"
         Top             =   4600
         Width           =   1000
      End
      Begin VB.TextBox pDc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1100
         TabIndex        =   15
         Text            =   "0"
         Top             =   4100
         Width           =   1500
      End
      Begin VB.TextBox pAc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1100
         TabIndex        =   14
         Text            =   "0"
         Top             =   3600
         Width           =   1500
      End
      Begin VB.ComboBox pnkpA 
         Height          =   405
         ItemData        =   "OZ.frx":032B
         Left            =   1100
         List            =   "OZ.frx":033E
         TabIndex        =   13
         Text            =   "1"
         Top             =   3100
         Width           =   800
      End
      Begin VB.TextBox phc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   700
         TabIndex        =   6
         Text            =   "0"
         Top             =   1900
         Width           =   1000
      End
      Begin VB.TextBox pYc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   700
         TabIndex        =   5
         Text            =   "0"
         Top             =   1400
         Width           =   1500
      End
      Begin VB.TextBox pXc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   700
         TabIndex        =   4
         Text            =   "0"
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label54 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ План. Цели"
         Height          =   300
         Left            =   2500
         TabIndex        =   162
         Top             =   2160
         Width           =   1700
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Глубина"
         Height          =   300
         Left            =   4300
         TabIndex        =   37
         Top             =   1400
         Width           =   1000
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фронт"
         Height          =   300
         Left            =   4300
         TabIndex        =   36
         Top             =   900
         Width           =   1000
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "    Размеры цели"
         Height          =   255
         Left            =   4300
         TabIndex        =   35
         Top             =   400
         Width           =   2200
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мц="
         Height          =   300
         Left            =   2900
         TabIndex        =   28
         Top             =   7300
         Width           =   500
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         Height          =   300
         Left            =   2900
         TabIndex        =   27
         Top             =   6800
         Width           =   500
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ КНП="
         Height          =   300
         Left            =   2900
         TabIndex        =   26
         Top             =   6300
         Width           =   1000
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мц="
         Height          =   300
         Left            =   100
         TabIndex        =   22
         Top             =   7300
         Width           =   500
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         Height          =   300
         Left            =   100
         TabIndex        =   21
         Top             =   6800
         Width           =   500
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ КНП="
         Height          =   300
         Left            =   100
         TabIndex        =   20
         Top             =   6300
         Width           =   1000
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "            ПРАВЫЙ"
         Height          =   300
         Left            =   2900
         TabIndex        =   19
         Top             =   5800
         Width           =   2500
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "             ЛЕВЫЙ"
         Height          =   300
         Left            =   100
         TabIndex        =   18
         Top             =   5800
         Width           =   2500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "                              СОПРЯЖЕНКА"
         Height          =   300
         Left            =   100
         TabIndex        =   17
         Top             =   5300
         Width           =   5300
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мц="
         Height          =   300
         Left            =   100
         TabIndex        =   12
         Top             =   4600
         Width           =   500
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Д="
         Height          =   300
         Left            =   100
         TabIndex        =   11
         Top             =   4100
         Width           =   500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         Height          =   300
         Left            =   100
         TabIndex        =   10
         Top             =   3600
         Width           =   500
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ КНП="
         Height          =   300
         Left            =   100
         TabIndex        =   9
         Top             =   3100
         Width           =   1000
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "                 А, Д"
         Height          =   300
         Left            =   100
         TabIndex        =   8
         Top             =   2600
         Width           =   2500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "             X, Y"
         Height          =   300
         Left            =   100
         TabIndex        =   7
         Top             =   400
         Width           =   2100
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   100
         TabIndex        =   3
         Top             =   1900
         Width           =   500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   100
         TabIndex        =   2
         Top             =   1400
         Width           =   500
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Х="
         Height          =   300
         Left            =   100
         TabIndex        =   1
         Top             =   900
         Width           =   500
      End
   End
End
Attribute VB_Name = "OZ"
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
Public Wx As Single, Wz As Single, Aatzo As Single, Fr As Single, Gl As Single, veer As Single, Sk As Single, nkpl As Single, nkpp As Single
Public snar As String, vzriv As String, stre As String, rep As String
Public vrv As Single, vsem As Single, Pricisch As Single, N As Single, dNtus As Single

Private Sub Command1_Click()
Dim Xkp As Single, Ykp As Single, hkp As Single, A1 As Single, A2 As Single, A3 As Single, A4 As Single, A5 As Single
Dim D1 As Single, D2 As Single, D3 As Single, D4 As Single, D5 As Single
Dim Mc1 As Single, Mc2 As Single, Mc3 As Single, Mc4 As Single, Mc5 As Single
Xkp = BP.pXkp1: Ykp = BP.pYkp1: hkp = BP.phkp1
YGLUKNP Dtkp, Yrkp, Ygoltkp, Xkp, Ykp, hkp
If Xkp = 0 Then
    A1 = 0: D1 = 0: Mc1 = 0
    Else
        A1 = Ygoltkp: D1 = Dtkp: Mc1 = Yrkp
End If
Xkp = BP.pXkp2: Ykp = BP.pYkp2: hkp = BP.phkp2
YGLUKNP Dtkp, Yrkp, Ygoltkp, Xkp, Ykp, hkp
If Xkp = 0 Then
    A2 = 0: D2 = 0: Mc2 = 0
    Else
        A2 = Ygoltkp: D2 = Dtkp: Mc2 = Yrkp
End If
Xkp = BP.pXkp3: Ykp = BP.pYkp3: hkp = BP.phkp3
YGLUKNP Dtkp, Yrkp, Ygoltkp, Xkp, Ykp, hkp
If Xkp = 0 Then
    A3 = 0: D3 = 0: Mc3 = 0
    Else
        A3 = Ygoltkp: D3 = Dtkp: Mc3 = Yrkp
End If
Xkp = BP.pXkp4: Ykp = BP.pYkp4: hkp = BP.phkp4
YGLUKNP Dtkp, Yrkp, Ygoltkp, Xkp, Ykp, hkp
If Xkp = 0 Then
    A4 = 0: D4 = 0: Mc4 = 0
    Else
        A4 = Ygoltkp: D4 = Dtkp: Mc4 = Yrkp
End If
Xkp = BP.pXkp5: Ykp = BP.pYkp5: hkp = BP.phkp5
YGLUKNP Dtkp, Yrkp, Ygoltkp, Xkp, Ykp, hkp
If Xkp = 0 Then
    A5 = 0: D5 = 0: Mc5 = 0
    Else
        A5 = Ygoltkp: D5 = Dtkp: Mc5 = Yrkp
End If
KNPpoZeli.Show
KNPpoZeli.pA1.Text = Round(A1): KNPpoZeli.pD1.Text = Round(D1): KNPpoZeli.pMc1.Text = Round(Mc1)
KNPpoZeli.pA2.Text = Round(A2): KNPpoZeli.pD2.Text = Round(D2): KNPpoZeli.pMc2.Text = Round(Mc2)
KNPpoZeli.pA3.Text = Round(A3): KNPpoZeli.pD3.Text = Round(D3): KNPpoZeli.pMc3.Text = Round(Mc3)
KNPpoZeli.pA4.Text = Round(A4): KNPpoZeli.pD4.Text = Round(D4): KNPpoZeli.pMc4.Text = Round(Mc4)
KNPpoZeli.pA5.Text = Round(A5): KNPpoZeli.pD5.Text = Round(D5): KNPpoZeli.pMc5.Text = Round(Mc5)
End Sub


Private Sub Command2_Click()
Spravka.Show
End Sub

Private Sub Command3_Click()
OZzelkagdform.Show
End Sub

Private Sub Command4_Click()
 OZ.Hide
End Sub

Private Sub Form_Load()
Dim t(1 To 10) As String
Dim i As Integer

941 Open "D:\YO_NA\zeli" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 942
 Input #1, t(1), t(2), t(3), t(4), t(5), t(6)
pplZel.AddItem t(1)
Loop
942 Close #1
End Sub

Private Sub okPopravki_Click()
otprkomandy.Show
End Sub

Private Sub OZAD_Click()
Dim Ac As Single, Dc As Single, Mc As Single, nkp As Single, Xkp As Single, Ykp As Single, hkp As Single, Xc As Single, Yc As Single, hc As Single
Dim dDov1 As Single, Dret1 As Single, dDr1 As Single, dN As Single
Dim rep1 As String, rep2 As String, rep3 As String

   Pi = 3.14159265358
Ac = pAc: Dc = pDc: Mc = pMc: nkp = pnkpA
If nkp = 1 Then
    Xkp = BP.pXkp1: Ykp = BP.pYkp1: hkp = BP.phkp1
    ElseIf nkp = 2 Then
    Xkp = BP.pXkp2: Ykp = BP.pYkp2: hkp = BP.phkp2
    ElseIf nkp = 3 Then
    Xkp = BP.pXkp3: Ykp = BP.pYkp3: hkp = BP.phkp3
    ElseIf nkp = 4 Then
    Xkp = BP.pXkp4: Ykp = BP.pYkp4: hkp = BP.phkp4
    Else
    Xkp = BP.pXkp5: Ykp = BP.pYkp5: hkp = BP.phkp5
End If
 Xc = Cos(Ac / 100 * 6 * Pi / 180) * Dc + Xkp
Yc = Sin(Ac / 100 * 6 * Pi / 180) * Dc + Ykp
hc = (Mc * (Dc * 0.001)) * 1.05 + hkp
pXc.Text = Int(Xc): pYc.Text = Int(Yc): phc.Text = Int(hc)

'записать номер цели в файл
Open App.Path & "\numberZeli" For Output As #1
Write #1, pplZel
Close #1

''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
50:
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
   If zo11 = 1 Then
   Xc = Xc1: Yc = Yc1: hc = hc
   Else
   Xc = pXc: Yc = pYc: hc = phc
   End If
   Xc1 = Xc: Yc1 = Yc: hc1 = hc
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = Xc - Xop1
60: dy1 = Yc - Yop1
61: dh1 = hc - hop1
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = pZar1
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

Dim snar As String

snar = pSnar1: vzriv = pVzr1
msgVelikaDalnost snar, zar, "1-я Батарея", Dt

       If stre = "Мортирная" Then
        podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
        podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       dddt1 = dddt: tz = tz1: zc1 = zc
       
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
        
       poddV0 tz, zar, dv0
       
       rep1 = pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
         If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt1 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt1 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt1 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt1 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   Else
                     popvnapk = dZwc * Wz + zc
                End If
            End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt + dXv0c * (v01 + dv0)
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
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
       podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr1: preps1 = CInt(Pric1 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr1: preps1 = CInt(Pric1 + daep)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
        Xc1 = Xc: Yc1 = Yc: hc1 = hc
        Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt1 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
    If BP.pX1 <> 0 Then
        pvSnar1.Text = snar: pvvzr1.Text = vzriv: pvZar1.Text = zar: pvPric1.Text = preps1: pvN1.Text = CInt(N1): pvDov1.Text = dovisch1
        pvVeer1.Text = veer: pvSk1.Text = Sk: pvdXtus1.Text = dXtus11: pvdNtus1.Text = dNtus1: pvPolet1.Text = ts1: pvVustra1.Text = Vustra1
        pvVd1.Text = Vd: pvDt1.Text = Dt1: pvYgt1.Text = Ygolt1: pvDovt1.Text = Dovort1: pvYr1.Text = Yr1: pvOH1.Text = OH1: pvdD1.Text = CInt(popvD)
        pvDisch1.Text = Int(Disch1): pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
    Else
            pvSnar1.Text = 0: pvvzr1.Text = 0: pvZar1.Text = 0: pvPric1.Text = 0: pvN1.Text = 0: pvDov1.Text = 0
        pvVeer1.Text = 0: pvSk1.Text = 0: pvdXtus1.Text = 0: pvdNtus1.Text = 0: pvPolet1.Text = 0: pvVustra1.Text = 0
        pvVd1.Text = 0: pvDt1.Text = 0: pvYgt1.Text = 0: pvDovt1.Text = 0: pvYr1.Text = 0: pvOH1.Text = 0: pvdD1.Text = 0
        pvDisch1.Text = 0: pvdDov1.Text = 0

End If
vrv = 0
 ' 2B
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
        If zo11 = 1 Then
         Xc = Xc2: Yc = Yc2: hc = hc
         Else
         Xc = Xc: Yc = Yc: hc = hc
         End If
         Xc2 = Xc: Yc2 = Yc: hc2 = hc
        dx2 = Xc - Xop2
104112:  dy2 = Yc - Yop2
104113:  dh2 = hc - hop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = pZar2
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
     
snar = pSnar2: vzriv = pVzr2
msgVelikaDalnost snar, zar, "2-я Батарея", Dt

       If stre = "Мортирная" Then
       podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        poddV0 tz, zar, dv0
       rep2 = pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt2 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt2 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt2 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt2 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
End If
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
      Xc2 = Xc: Yc2 = Yc: hc2 = hc
              Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt2 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
    If BP.pX2 <> 0 Then
              pvSnar2.Text = snar: pvvzr2.Text = vzriv: pvZar2.Text = zar: pvPric2.Text = preps2: pvN2.Text = CInt(N2): pvDov2.Text = dovisch2
        pvVeer2.Text = veer: pvSk2.Text = Sk: pvdXtus2.Text = dXtus2: pvdNtus2.Text = dNtus2: pvPolet2.Text = ts2: pvVustra2.Text = Vustra2
        pvVd2.Text = Vd: pvDt2.Text = Dt2: pvYgt2.Text = Ygolt2: pvDovt2.Text = Dovort2: pvYr2.Text = Yr2: pvOH2.Text = OH2: pvdD2.Text = CInt(popvD)
        pvDisch2.Text = Int(Disch2): pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
              pvSnar2.Text = 0: pvvzr2.Text = 0: pvZar2.Text = 0: pvPric2.Text = 0: pvN2.Text = 0: pvDov2.Text = 0
        pvVeer2.Text = 0: pvSk2.Text = 0: pvdXtus2.Text = 0: pvdNtus2.Text = 0: pvPolet2.Text = 0: pvVustra2.Text = 0
        pvVd2.Text = 0: pvDt2.Text = 0: pvYgt2.Text = 0: pvDovt2.Text = 0: pvYr2.Text = 0: pvOH2.Text = 0: pvdD2.Text = 0
        pvDisch2.Text = 0: pvdDov2.Text = 0
End If
vrv = 0
  '3B
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
        If zo11 = 1 Then
          Xc = Xc3: Yc = Yc3: hc = hc
          Else
          Xc = Xc: Yc = Yc: hc = hc
          End If
          Xc3 = Xc: Yc3 = Yc: hc3 = hc
         dx3 = Xc - Xop3
1041194:  dy3 = Yc - Yop3
1041195:  dh3 = hc - hop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = pZar3
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
       
snar = pSnar3: vzriv = pVzr3
msgVelikaDalnost snar, zar, "3-я Батарея", Dt

If stre = "Мортирная" Then
       podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     poddV0 tz, zar, dv0
            rep3 = pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
    If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt3 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt3 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt3 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt3 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
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
        podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
       Xc3 = Xc: Yc3 = Yc: hc3 = hc
               Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt3 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX3 <> 0 Then
                     pvSnar3.Text = snar: pvvzr3.Text = vzriv: pvZar3.Text = zar: pvPric3.Text = preps3: pvN3.Text = CInt(N3): pvDov3.Text = dovisch3
        pvVeer3.Text = veer: pvSk3.Text = Sk: pvdXtus3.Text = dXtus3: pvdNtus3.Text = dNtus3: pvPolet3.Text = ts3: pvVustra3.Text = Vustra3
        pvVd3.Text = Vd: pvDt3.Text = Dt3: pvYgt3.Text = Ygolt3: pvDovt3.Text = Dovort3: pvYr3.Text = Yr3: pvOH3.Text = OH3: pvdD3.Text = CInt(popvD)
        pvDisch3.Text = Int(Disch3): pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
                     pvSnar3.Text = 0: pvvzr3.Text = 0: pvZar3.Text = 0: pvPric3.Text = 0: pvN3.Text = 0: pvDov3.Text = 0
        pvVeer3.Text = 0: pvSk3.Text = 0: pvdXtus3.Text = 0: pvdNtus3.Text = 0: pvPolet3.Text = 0: pvVustra3.Text = 0
        pvVd3.Text = 0: pvDt3.Text = 0: pvYgt3.Text = 0: pvDovt3.Text = 0: pvYr3.Text = 0: pvOH3.Text = 0: pvdD3.Text = 0
        pvDisch3.Text = 0: pvdDov3.Text = 0
End If
vrv = 0
End Sub

Private Sub OZSopr_Click()
'записать номер цели в файл
Open App.Path & "\numberZeli" For Output As #1
Write #1, pplZel
Close #1

 Dim dDov1 As Single, Dret1 As Single, dDr1 As Single, dN As Single
 Dim rep1 As String, rep2 As String, rep3 As String
nkpl = pNKPL: nkpp = pNKPP: Alev = pAlc: Aprav = pApc: Mcl = pMlc: Mcp = pMpc
If nkpl = 1 Then
    Xkpl = BP.pXkp1: Ykpl = BP.pYkp1: hkpl = BP.phkp1
    ElseIf nkpl = 2 Then
        Xkpl = BP.pXkp2: Ykpl = BP.pYkp2: hkpl = BP.phkp2
        ElseIf nkpl = 3 Then
            Xkpl = BP.pXkp3: Ykpl = BP.pYkp3: hkpl = BP.phkp3
            ElseIf nkpl = 4 Then
                Xkpl = BP.pXkp4: Ykpl = BP.pYkp4: hkpl = BP.phkp4
                Else
                    Xkpl = BP.pXkp5: Ykpl = BP.pYkp5: hkpl = BP.phkp5
End If
If nkpp = 1 Then
    Xkpp = BP.pXkp1: Ykpp = BP.pYkp1: hkpp = BP.phkp1
    ElseIf nkpp = 2 Then
        Xkpp = BP.pXkp2: Ykpp = BP.pYkp2: hkpp = BP.phkp2
        ElseIf nkpp = 3 Then
            Xkpp = BP.pXkp3: Ykpp = BP.pYkp3: hkpp = BP.phkp3
            ElseIf nkpp = 4 Then
                Xkpp = BP.pXkp4: Ykpp = BP.pYkp4: hkpp = BP.phkp4
                Else
                    Xkpp = BP.pXkp5: Ykpp = BP.pYkp5: hkpp = BP.phkp5
End If
 dxso = Xkpp - Xkpl: dyso = Ykpp - Ykpl
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
  Xcso = Cos(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Xkpl
  Ycso = Sin(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Ykpl
  If Mcl = 0 Then hc = Mcp * (Dprav * 0.001) * 1.05 + hkpp
  If Mcp = 0 Then hc = Mcl * (Dlev * 0.001) * 1.05 + hkpl
  Xc = Xcso: Yc = Ycso
  pXc.Text = Int(Xc): pYc.Text = Int(Yc): phc.Text = Int(hc)

   '1B
50:
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
   If zo11 = 1 Then
   Xc = Xc1: Yc = Yc1: hc = hc
   Else
   Xc = pXc: Yc = pYc: hc = phc
   End If
   Xc1 = Xc: Yc1 = Yc: hc1 = hc
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = Xc - Xop1
60: dy1 = Yc - Yop1
61: dh1 = hc - hop1
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = pZar1
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

snar = pSnar1: vzriv = pVzr1
msgVelikaDalnost snar, zar, "1-я Батарея", Dt

       If stre = "Мортирная" Then
       podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       poddV0 tz, zar, dv0
              rep1 = pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
   If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt1 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt1 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
             If popvD < 0 Then
                   Dt = Dt1 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt1 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   Else
                     popvnapk = dZwc * Wz + zc
                End If
            End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt + dXv0c * (v01 + dv0)
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
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
       podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr1: preps1 = CInt(Pric1 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr1: preps1 = CInt(Pric1 + daep)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
        Xc1 = Xc: Yc1 = Yc: hc1 = hc
                Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt1 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX1 <> 0 Then
        pvSnar1.Text = snar: pvvzr1.Text = vzriv: pvZar1.Text = zar: pvPric1.Text = preps1: pvN1.Text = CInt(N1): pvDov1.Text = dovisch1
        pvVeer1.Text = veer: pvSk1.Text = Sk: pvdXtus1.Text = dXtus11: pvdNtus1.Text = dNtus1: pvPolet1.Text = ts1: pvVustra1.Text = Vustra1
        pvVd1.Text = Vd: pvDt1.Text = Dt1: pvYgt1.Text = Ygolt1: pvDovt1.Text = Dovort1: pvYr1.Text = Yr1: pvOH1.Text = OH1: pvdD1.Text = CInt(popvD)
        pvDisch1.Text = Int(Disch1): pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
            pvSnar1.Text = 0: pvvzr1.Text = 0: pvZar1.Text = 0: pvPric1.Text = 0: pvN1.Text = 0: pvDov1.Text = 0
        pvVeer1.Text = 0: pvSk1.Text = 0: pvdXtus1.Text = 0: pvdNtus1.Text = 0: pvPolet1.Text = 0: pvVustra1.Text = 0
        pvVd1.Text = 0: pvDt1.Text = 0: pvYgt1.Text = 0: pvDovt1.Text = 0: pvYr1.Text = 0: pvOH1.Text = 0: pvdD1.Text = 0
        pvDisch1.Text = 0: pvdDov1.Text = 0

End If
vrv = 0
 ' 2B
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
        If zo11 = 1 Then
         Xc = Xc2: Yc = Yc2: hc = hc
         Else
         Xc = Xc: Yc = Yc: hc = hc
         End If
         Xc2 = Xc: Yc2 = Yc: hc2 = hc
        dx2 = Xc - Xop2
104112:  dy2 = Yc - Yop2
104113:  dh2 = hc - hop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = pZar2
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
     
snar = pSnar2: vzriv = pVzr2
msgVelikaDalnost snar, zar, "2-я Батарея", Dt

       If stre = "Мортирная" Then
       podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        poddV0 tz, zar, dv0
               rep2 = pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
     If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt2 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt2 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt2 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt2 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
End If
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
      Xc2 = Xc: Yc2 = Yc: hc2 = hc
              Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt2 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX2 <> 0 Then
              pvSnar2.Text = snar: pvvzr2.Text = vzriv: pvZar2.Text = zar: pvPric2.Text = preps2: pvN2.Text = CInt(N2): pvDov2.Text = dovisch2
        pvVeer2.Text = veer: pvSk2.Text = Sk: pvdXtus2.Text = dXtus2: pvdNtus2.Text = dNtus2: pvPolet2.Text = ts2: pvVustra2.Text = Vustra2
        pvVd2.Text = Vd: pvDt2.Text = Dt2: pvYgt2.Text = Ygolt2: pvDovt2.Text = Dovort2: pvYr2.Text = Yr2: pvOH2.Text = OH2: pvdD2.Text = CInt(popvD)
        pvDisch2.Text = Int(Disch2): pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
              pvSnar2.Text = 0: pvvzr2.Text = 0: pvZar2.Text = 0: pvPric2.Text = 0: pvN2.Text = 0: pvDov2.Text = 0
        pvVeer2.Text = 0: pvSk2.Text = 0: pvdXtus2.Text = 0: pvdNtus2.Text = 0: pvPolet2.Text = 0: pvVustra2.Text = 0
        pvVd2.Text = 0: pvDt2.Text = 0: pvYgt2.Text = 0: pvDovt2.Text = 0: pvYr2.Text = 0: pvOH2.Text = 0: pvdD2.Text = 0
        pvDisch2.Text = 0: pvdDov2.Text = 0
End If
vrv = 0
  '3B
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
        If zo11 = 1 Then
          Xc = Xc3: Yc = Yc3: hc = hc
          Else
          Xc = Xc: Yc = Yc: hc = hc
          End If
          Xc3 = Xc: Yc3 = Yc: hc3 = hc
         dx3 = Xc - Xop3
1041194:  dy3 = Yc - Yop3
1041195:  dh3 = hc - hop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = pZar3
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

snar = pSnar3: vzriv = pVzr3
msgVelikaDalnost snar, zar, "3-я Батарея", Dt

       If stre = "Мортирная" Then
       podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     poddV0 tz, zar, dv0
            rep3 = pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
      If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt3 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt3 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt3 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt3 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
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
        podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
       Xc3 = Xc: Yc3 = Yc: hc3 = hc
               Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt3 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX3 <> 0 Then
                     pvSnar3.Text = snar: pvvzr3.Text = vzriv: pvZar3.Text = zar: pvPric3.Text = preps3: pvN3.Text = CInt(N3): pvDov3.Text = dovisch3
        pvVeer3.Text = veer: pvSk3.Text = Sk: pvdXtus3.Text = dXtus3: pvdNtus3.Text = dNtus3: pvPolet3.Text = ts3: pvVustra3.Text = Vustra3
        pvVd3.Text = Vd: pvDt3.Text = Dt3: pvYgt3.Text = Ygolt3: pvDovt3.Text = Dovort3: pvYr3.Text = Yr3: pvOH3.Text = OH3: pvdD3.Text = CInt(popvD)
        pvDisch3.Text = Int(Disch3): pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
                     pvSnar3.Text = 0: pvvzr3.Text = 0: pvZar3.Text = 0: pvPric3.Text = 0: pvN3.Text = 0: pvDov3.Text = 0
        pvVeer3.Text = 0: pvSk3.Text = 0: pvdXtus3.Text = 0: pvdNtus3.Text = 0: pvPolet3.Text = 0: pvVustra3.Text = 0
        pvVd3.Text = 0: pvDt3.Text = 0: pvYgt3.Text = 0: pvDovt3.Text = 0: pvYr3.Text = 0: pvOH3.Text = 0: pvdD3.Text = 0
        pvDisch3.Text = 0: pvdDov3.Text = 0
End If
vrv = 0
End Sub

Private Sub OZXY_Click()
 Dim dDov1 As Single, Dret1 As Single, dDr1 As Single, dN As Single
 Dim rep1 As String, rep2 As String, rep3 As String
 Dim A1 As Single, ras As Single, dhh1 As Single, zo11 As Single
 Dim Xc1 As Single, Yc1 As Single, hc1 As Single, dx1 As Single, dy1 As Single
 Dim dh1 As Single, Pi As Single, Dt As Single, Ygolt As Single, dh As Single
 Dim v01 As Single, dXtus As Single, ybyl As Single, ts As Single, zc As Single
 Dim dZwc As Single, dXwc As Single, dXhc As Single, dXtc As Single, dXv0c As Single
 Dim met As Single, ybylc As Single, dddt As Single, Ygvozv As Single, Ygpad As Single
Dim Vustra As Single, Vd As Single, dddt1 As Single, tz As Single, zc1 As Single
Dim tsk As Single, dXtusk As Single, Ygvozvk As Single, Vustrak As Single, Ygpadk As Single
Dim Vdk As Single, dv0 As Single, popvnap As Single, popvD As Single, Dtk As Single
Dim popvnapk As Single, popvdk As Single, Dtisch As Single, Dtischk As Single
Dim kPop As Single, popvd1 As Single, Disch As Single, Kpopnap As Single, popvnap1 As Single
Dim dovisch1 As Single, dhh As Single, dV00 As Single, Yr As Single, Yrr As Single
Dim Ygpad1 As Single, Ygvozv1 As Single, Vustra1 As Single, dXtus11 As Single
Dim kpe As Single, daep As Single, preps1 As Single
Dim zar As String

'записать номер цели в файл
Open App.Path & "\numberZeli" For Output As #1
Write #1, pplZel
Close #1

''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
50:
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
   If zo11 = 1 Then
   Xc = Xc1: Yc = Yc1: hc = hc
   Else
   Xc = pXc: Yc = pYc: hc = phc
   End If
   Xc1 = Xc: Yc1 = Yc: hc1 = hc
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = Xc - Xop1
60: dy1 = Yc - Yop1
61: dh1 = hc - hop1
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = pZar1
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

snar = pSnar1: vzriv = pVzr1
msgVelikaDalnost snar, zar, "1-я Батарея", Dt

       If stre = "Мортирная" Then
       podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       poddV0 tz, zar, dv0
              rep1 = pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt1 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt1 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
                 If popvD < 0 Then
                   Dt = Dt1 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt1 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   Else
                     popvnapk = dZwc * Wz + zc
                End If
        End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt + dXv0c * (v01 + dv0)
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
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
       podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr1: preps1 = CInt(Pric1 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr1: preps1 = CInt(Pric1 + daep)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
        Xc1 = Xc: Yc1 = Yc: hc1 = hc
                Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt1 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX1 <> 0 Then
        pvSnar1.Text = snar: pvvzr1.Text = vzriv: pvZar1.Text = zar: pvPric1.Text = preps1: pvN1.Text = CInt(N1): pvDov1.Text = dovisch1
        pvVeer1.Text = veer: pvSk1.Text = Sk: pvdXtus1.Text = dXtus11: pvdNtus1.Text = dNtus1: pvPolet1.Text = ts1: pvVustra1.Text = Vustra1
        pvVd1.Text = Vd: pvDt1.Text = Dt1: pvYgt1.Text = Ygolt1: pvDovt1.Text = Dovort1: pvYr1.Text = Yr1: pvOH1.Text = OH1: pvdD1.Text = CInt(popvD)
        pvDisch1.Text = Int(Disch1): pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
            pvSnar1.Text = 0: pvvzr1.Text = 0: pvZar1.Text = 0: pvPric1.Text = 0: pvN1.Text = 0: pvDov1.Text = 0
        pvVeer1.Text = 0: pvSk1.Text = 0: pvdXtus1.Text = 0: pvdNtus1.Text = 0: pvPolet1.Text = 0: pvVustra1.Text = 0
        pvVd1.Text = 0: pvDt1.Text = 0: pvYgt1.Text = 0: pvDovt1.Text = 0: pvYr1.Text = 0: pvOH1.Text = 0: pvdD1.Text = 0
        pvDisch1.Text = 0: pvdDov1.Text = 0
End If
vrv = 0
 ' 2B
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
        If zo11 = 1 Then
         Xc = Xc2: Yc = Yc2: hc = hc
         Else
         Xc = Xc: Yc = Yc: hc = hc
         End If
         Xc2 = Xc: Yc2 = Yc: hc2 = hc
        dx2 = Xc - Xop2
104112:  dy2 = Yc - Yop2
104113:  dh2 = hc - hop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = pZar2
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
     
snar = pSnar2: vzriv = pVzr2
msgVelikaDalnost snar, zar, "2-я Батарея", Dt

       If stre = "Мортирная" Then
       podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        poddV0 tz, zar, dv0
               rep2 = pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt2 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt2 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt2 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt2 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
End If
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
      Xc2 = Xc: Yc2 = Yc: hc2 = hc
              Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt2 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX2 <> 0 Then
              pvSnar2.Text = snar: pvvzr2.Text = vzriv: pvZar2.Text = zar: pvPric2.Text = preps2: pvN2.Text = CInt(N2): pvDov2.Text = dovisch2
        pvVeer2.Text = veer: pvSk2.Text = Sk: pvdXtus2.Text = dXtus2: pvdNtus2.Text = dNtus2: pvPolet2.Text = ts2: pvVustra2.Text = Vustra2
        pvVd2.Text = Vd: pvDt2.Text = Dt2: pvYgt2.Text = Ygolt2: pvDovt2.Text = Dovort2: pvYr2.Text = Yr2: pvOH2.Text = OH2: pvdD2.Text = CInt(popvD)
        pvDisch2.Text = Int(Disch2): pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
              pvSnar2.Text = 0: pvvzr2.Text = 0: pvZar2.Text = 0: pvPric2.Text = 0: pvN2.Text = 0: pvDov2.Text = 0
        pvVeer2.Text = 0: pvSk2.Text = 0: pvdXtus2.Text = 0: pvdNtus2.Text = 0: pvPolet2.Text = 0: pvVustra2.Text = 0
        pvVd2.Text = 0: pvDt2.Text = 0: pvYgt2.Text = 0: pvDovt2.Text = 0: pvYr2.Text = 0: pvOH2.Text = 0: pvdD2.Text = 0
        pvDisch2.Text = 0: pvdDov2.Text = 0
End If
vrv = 0
  '3B
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
        If zo11 = 1 Then
          Xc = Xc3: Yc = Yc3: hc = hc
          Else
          Xc = Xc: Yc = Yc: hc = hc
          End If
          Xc3 = Xc: Yc3 = Yc: hc3 = hc
         dx3 = Xc - Xop3
1041194:  dy3 = Yc - Yop3
1041195:  dh3 = hc - hop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = pZar3
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
       
snar = pSnar3: vzriv = pVzr3
msgVelikaDalnost snar, zar, "3-я Батарея", Dt

       If stre = "Мортирная" Then
       podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     poddV0 tz, zar, dv0
            rep3 = pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
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
                podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        If popvD < 0 And stre = "Мортирная" Then
            Dt = Dt3 - 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            popvnapk = dZwc * Wz + zc
            Dt = Dt3 + 1000
            podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
            Else
              If popvD < 0 Then
                   Dt = Dt3 - 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
                   popvnapk = dZwc * Wz + zc
                   Dt = Dt3 + 1000
                   podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
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
      podPRICMORTRGM zar, Disch, Pricisch, ts
      ElseIf vzriv = "АР-5" Then
      podAR5 zar, Disch, Pricisch, N
      ElseIf vzriv = "ДТМ-75" Then
      pod3SH1 Disch, zar, rep, vsem, Pricisch, N, dNtus
      ElseIf vzriv = "В-90" Then
      podB90 zar, Disch, rep, Wx, N, dNtus, vrv, Pricisch
      ElseIf vzriv = "Т-90" Then
      podT90 Disch, zar, N, dNtus, Pricisch
      Else
      podPRICRGM zar, snar, Disch, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd
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
        podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
       Xc3 = Xc: Yc3 = Yc: hc3 = hc
               Fr = pFrontc: Gl = pGlybinac
        veer = Int(Fr / ((Dt3 + 0.001) / 1000) * 0.95)
        Sk = Int((Gl + 0.001) / 3 / (dXtus + 0.001))
If BP.pX3 <> 0 Then
                     pvSnar3.Text = snar: pvvzr3.Text = vzriv: pvZar3.Text = zar: pvPric3.Text = preps3: pvN3.Text = CInt(N3): pvDov3.Text = dovisch3
        pvVeer3.Text = veer: pvSk3.Text = Sk: pvdXtus3.Text = dXtus3: pvdNtus3.Text = dNtus3: pvPolet3.Text = ts3: pvVustra3.Text = Vustra3
        pvVd3.Text = Vd: pvDt3.Text = Dt3: pvYgt3.Text = Ygolt3: pvDovt3.Text = Dovort3: pvYr3.Text = Yr3: pvOH3.Text = OH3: pvdD3.Text = CInt(popvD)
        pvDisch3.Text = Int(Disch3): pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
                     pvSnar3.Text = 0: pvvzr3.Text = 0: pvZar3.Text = 0: pvPric3.Text = 0: pvN3.Text = 0: pvDov3.Text = 0
        pvVeer3.Text = 0: pvSk3.Text = 0: pvdXtus3.Text = 0: pvdNtus3.Text = 0: pvPolet3.Text = 0: pvVustra3.Text = 0
        pvVd3.Text = 0: pvDt3.Text = 0: pvYgt3.Text = 0: pvDovt3.Text = 0: pvYr3.Text = 0: pvOH3.Text = 0: pvdD3.Text = 0
        pvDisch3.Text = 0: pvdDov3.Text = 0
End If
vrv = 0
  End Sub
''''''''''''''''''''''''''POPRAVKI''''''''''''''''''''''''''''''''''''''''''''
Function podPOPRAVKI(ByVal zar As String, ByVal snar As String, ByVal Dt As Single, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd) As Single
Dim ta(1 To 50) As Single
Dim Dta As Single, Prta As Single, z As Single, dzw As Single, dxw As Single, dxh As Single
Dim dxt As Single, dxv0 As Single, ybyl As Single, dtT As Single, dta1 As Single
Dim prta1 As Single, z1 As Single, dzw1 As Single, dxw1 As Single, dxh1 As Single
Dim dxt1 As Single, dxv01 As Single, ybyl1 As Single, Aw200 As Single, Aw400 As Single
Dim Aw800 As Single, Aw1200 As Single, Aw1600 As Single, Aw2000 As Single, Aw2400 As Single
Dim Aw3000 As Single, Aw4000 As Single, Aw5000 As Single, Aw6000 As Single
Dim ybylc As Single, aww As Single, aws As Single, sw As Single
   
If snar = "3ОФ56" Then
    If zar = "Полн" Then
        Open App.Path & "\OF56Pol.txt" For Input As #1
        ElseIf zar = "Умен" Then
        Open App.Path & "\OF56Ymen.txt" For Input As #1
        ElseIf zar = "Перв" Then
        Open App.Path & "\OF56Perv.txt" For Input As #1
        ElseIf zar = "Втор" Then
        Open App.Path & "\OF56Vtor.txt" For Input As #1
        ElseIf zar = "Трет" Then
        Open App.Path & "\OF56Tret.txt" For Input As #1
        ElseIf zar = "Четверт" Then
        Open App.Path & "\OF56Chetvert.txt" For Input As #1
        Else
        Open App.Path & "\OF56Pol.txt" For Input As #1
    End If
    Else
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
End If
91: If EOF(1) Then GoTo 9212
9111: Input #1, ta(1), ta(2), ta(3), ta(4), ta(5), ta(6), ta(7), ta(8), ta(9), ta(10), ta(11), ta(12), ta(13), ta(14), ta(15), ta(16), ta(17), ta(18), ta(19), ta(20), ta(21), ta(22), ta(23), ta(24), ta(25), ta(26), ta(27), ta(28), ta(29), ta(30), ta(31), ta(32), ta(33), ta(34), ta(35), ta(36), ta(37)
92: If ta(1) <= Dt Or EOF(1) Then GoTo 9211
 GoTo 91
9211: Dta = ta(1): Prta = ta(2): dXtus = ta(3): z = ta(6): dzw = ta(7): dxw = ta(8)
Ygvozv = ta(13): Vd = ta(4): dxh = ta(9): dxt = ta(11): dxv0 = ta(12): ybyl = ta(18)
ts = ta(16): Ygpad = ta(14): Vustra = ta(17)
9212: Close #1
   dtT = Dt + 200
   
If snar = "3ОФ56" Then
    If zar = "Полн" Then
        Open App.Path & "\OF56Pol.txt" For Input As #1
        ElseIf zar = "Умен" Then
        Open App.Path & "\OF56Ymen.txt" For Input As #1
        ElseIf zar = "Перв" Then
        Open App.Path & "\OF56Perv.txt" For Input As #1
        ElseIf zar = "Втор" Then
        Open App.Path & "\OF56Vtor.txt" For Input As #1
        ElseIf zar = "Трет" Then
        Open App.Path & "\OF56Tret.txt" For Input As #1
        ElseIf zar = "Четверт" Then
        Open App.Path & "\OF56Chetvert.txt" For Input As #1
        Else
        Open App.Path & "\OF56Pol.txt" For Input As #1
    End If
    Else
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
End If

911: If EOF(1) Then GoTo 92121
91111: Input #1, ta(1), ta(2), ta(3), ta(4), ta(5), ta(6), ta(7), ta(8), ta(9), ta(10), ta(11), ta(12), ta(13), ta(14), ta(15), ta(16), ta(17), ta(18), ta(19), ta(20), ta(21), ta(22), ta(23), ta(24), ta(25), ta(26), ta(27), ta(28), ta(29), ta(30), ta(31), ta(32), ta(33), ta(34), ta(35), ta(36), ta(37)
921: If ta(1) <= dtT Or EOF(1) Then GoTo 92111
 GoTo 911
92111: dta1 = ta(1): prta1 = ta(2): dXtus1 = ta(3): z1 = ta(6): dzw1 = ta(7)
dxw1 = ta(8): dxh1 = ta(9): dxt1 = ta(11): dxv01 = ta(12): ybyl1 = ta(18)
92121: Close #1
zc = ((z1 - z) / 200 * (Dt - Dta) + z) * -1
dZwc = ((dzw1 - dzw) / 200 * (Dt - Dta) + dzw) * -1 / 10
dXwc = ((dxw1 - dxw) / 200 * (Dt - Dta) + dxw) * -1 / 10
dXhc = ((dxh1 - dxh) / 200 * (Dt - Dta) + dxh) / 10
dXtc = ((dxt1 - dxt) / 200 * (Dt - Dta) + dxt) * -1 / 10
dXv0c = ((dxv01 - dxv0) / 200 * (Dt - Dta) + dxv0) * -1
 Aw200 = Val(BP.pAw02): Aw400 = Val(BP.pAw04): Aw800 = Val(BP.pAw08)
 Aw1200 = Val(BP.pAw12): Aw1600 = Val(BP.pAw16): Aw2000 = Val(BP.pAw20)
 Aw2400 = Val(BP.pAw24): Aw3000 = Val(BP.pAw30): Aw4000 = Val(BP.pAw40)
 Aw5000 = Val(BP.pAw50): Aw6000 = Val(BP.pAw60)
ybylc = ybyl * 100
 If ybylc = 200 Then aww = Aw200
 If ybylc = 400 Then aww = Aw400
 If ybylc = 800 Then aww = Aw800
 If ybylc = 1200 Then aww = Aw1200
 If ybylc = 1600 Then aww = Aw1600
 If ybylc = 2000 Then aww = Aw2000
 If ybylc = 2400 Then aww = Aw2400
 If ybylc = 3000 Then aww = Aw3000
 If ybylc = 4000 Then aww = Aw4000
  If ybylc = 5000 Then aww = Aw5000
 If ybylc = 6000 Then aww = Aw6000
 aww = aww * 100
If Ygolt - aww < 0 Then
       aws = Ygolt + 6000 - aww
   Else
   aws = Ygolt - aww
   End If
 If ybylc = 200 Then sw = BP.pW02
 If ybylc = 400 Then sw = BP.pW04
 If ybylc = 800 Then sw = BP.pW08
 If ybylc = 1200 Then sw = BP.pW12
 If ybylc = 1600 Then sw = BP.pW16
 If ybylc = 2000 Then sw = BP.pW20
 If ybylc = 2400 Then sw = BP.pW24
 If ybylc = 3000 Then sw = BP.pW30
 If ybylc = 4000 Then sw = BP.pW40
  If ybylc = 5000 Then sw = BP.pW50
 If ybylc = 6000 Then sw = BP.pW60
 Wx = Abs(sw * Cos((aws / 100 * 6) * 3.141596 / 180))
 Wz = Abs(sw * Sin((aws / 100 * 6) * 3.14159 / 180))

 If aws <= 1500 Or aws >= 4500 Then Wx = Wx * -1
 If aws >= 3000 Then Wz = Wz * -1

 If ybylc = 200 Then dddt = BP.pdT02
 If ybylc = 400 Then dddt = BP.pdT04
 If ybylc = 800 Then dddt = BP.pdT08
 If ybylc = 1200 Then dddt = BP.pdT12
 If ybylc = 1600 Then dddt = BP.pdT16
 If ybylc = 2000 Then dddt = BP.pdT20
 If ybylc = 2400 Then dddt = BP.pdT24
 If ybylc = 3000 Then dddt = BP.pdT30
 If ybylc = 4000 Then dddt = BP.pdT40
 If ybylc = 5000 Then dddt = BP.pdT50
 If ybylc = 6000 Then dddt = BP.pDt60
 If dddt >= 50 Then dddt = (dddt - 50) * -1
  End Function
'''''''''''''''''''''PRICEL''''''''''''''''''''''''''''''''''''''''
Function podPRICRGM(ByVal zar As String, ByVal snar As String, ByVal Disch As Single, Pricisch, ts, dXtus, Ygvozv, Vustra, Vd) As Single
Dim ta(1 To 50) As Single
Dim Pric As Single, Dta As Single, Dischh As Single, Pricc As Single

If snar = "3ОФ56" Then
    If zar = "Полн" Then
            Open App.Path & "\OF56Pol.txt" For Input As #1
        ElseIf zar = "Умен" Then
            Open App.Path & "\OF56Ymen.txt" For Input As #1
        ElseIf zar = "Перв" Then
            Open App.Path & "\OF56Perv.txt" For Input As #1
        ElseIf zar = "Втор" Then
            Open App.Path & "\OF56Vtor.txt" For Input As #1
        ElseIf zar = "Трет" Then
            Open App.Path & "\OF56Tret.txt" For Input As #1
        ElseIf zar = "Четверт" Then
            Open App.Path & "\OF56Chetvert.txt" For Input As #1
        Else
            Open App.Path & "\OF56Pol.txt" For Input As #1
    End If
    ElseIf snar = "ОФ99" Then
        If zar = "Полн" Then
                Open App.Path & "\OF99\2C1P" For Input As #1
            ElseIf zar = "Умен" Then
                Open App.Path & "\OF99\2C1Y" For Input As #1
            ElseIf zar = "Перв" Then
                Open App.Path & "\OF99\2C11" For Input As #1
            ElseIf zar = "Втор" Then
                Open App.Path & "\OF99\2C12" For Input As #1
            ElseIf zar = "Трет" Then
                Open App.Path & "\OF99\2C13" For Input As #1
            ElseIf zar = "Четверт" Then
                Open App.Path & "\OF99\2C14" For Input As #1
            Else
                Open App.Path & "\OF99\2C1P" For Input As #1
            End If
    Else
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
End If

92131: If EOF(1) Then GoTo 92132
  Input #1, ta(1), ta(2), ta(3), ta(4), ta(5), ta(6), ta(7), ta(8), ta(9), ta(10), ta(11), ta(12), ta(13), ta(14), ta(15), ta(16), ta(17), ta(18), ta(19), ta(20), ta(21), ta(22), ta(23), ta(24), ta(25), ta(26), ta(27), ta(28), ta(29), ta(30), ta(31), ta(32), ta(33), ta(34), ta(35), ta(36), ta(37)
  If ta(1) <= Disch Or EOF(1) Then GoTo 921311
      GoTo 92131
921311: Pric = ta(2): Dta = ta(1): ts = ta(16): dXtus = ta(3): Ygvozv = ta(13)
        Vustra = ta(17): Vd = ta(4)
92132: Close #1
  Dischh = Disch + 200

If snar = "3ОФ56" Then
    If zar = "Полн" Then
            Open App.Path & "\OF56Pol.txt" For Input As #1
        ElseIf zar = "Умен" Then
            Open App.Path & "\OF56Ymen.txt" For Input As #1
        ElseIf zar = "Перв" Then
            Open App.Path & "\OF56Perv.txt" For Input As #1
        ElseIf zar = "Втор" Then
            Open App.Path & "\OF56Vtor.txt" For Input As #1
        ElseIf zar = "Трет" Then
            Open App.Path & "\OF56Tret.txt" For Input As #1
        ElseIf zar = "Четверт" Then
            Open App.Path & "\OF56Chetvert.txt" For Input As #1
        Else
            Open App.Path & "\OF56Pol.txt" For Input As #1
    End If
    ElseIf snar = "ОФ99" Then
        If zar = "Полн" Then
                Open App.Path & "\OF99\2C1P" For Input As #1
            ElseIf zar = "Умен" Then
                Open App.Path & "\OF99\2C1Y" For Input As #1
            ElseIf zar = "Перв" Then
                Open App.Path & "\OF99\2C11" For Input As #1
            ElseIf zar = "Втор" Then
                Open App.Path & "\OF99\2C12" For Input As #1
            ElseIf zar = "Трет" Then
                Open App.Path & "\OF99\2C13" For Input As #1
            ElseIf zar = "Четверт" Then
                Open App.Path & "\OF99\2C14" For Input As #1
            Else
                Open App.Path & "\OF99\2C1P" For Input As #1
            End If
    Else
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
End If

921321: If EOF(1) Then GoTo 921322
     Input #1, ta(1), ta(2), ta(3), ta(4), ta(5), ta(6), ta(7), ta(8), ta(9), ta(10), ta(11), ta(12), ta(13), ta(14), ta(15), ta(16), ta(17), ta(18), ta(19), ta(20), ta(21), ta(22), ta(23), ta(24), ta(25), ta(26), ta(27), ta(28), ta(29), ta(30), ta(31), ta(32), ta(33), ta(34), ta(35), ta(36), ta(37)
   If ta(1) <= Dischh Or EOF(1) Then GoTo 9213211
   GoTo 921321
9213211: Pricc = ta(2)
921322:  Close #1
    Pricisch = (Pricc - Pric) / 200 * (Disch - Dta) + Pric
End Function
''''''''''''''''''''''''''''''KPE'''''''''''''''''''''''''''''''''''''''''
Function podKPE(ByVal zar As String, ByVal Pricisch As Single, ByVal Yrr As Single, kpe) As Single
Dim e(1 To 10) As Single
Dim kpew As Single, kpen As Single

223:  If zar = "Полн" Then
        Open App.Path & "\2c1kpep" For Input As 1
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
22311: If EOF(1) Then GoTo 22312
  Input #1, e(1), e(2), e(3), e(4), e(5)
   If Pricisch >= e(1) And e(1) + 20 >= Pricisch Then kpew = e(2): kpen = e(3): GoTo 22312
   GoTo 22311
22312: Close #1
    If Yrr >= 0 Then kpe = kpew * 0.1
    If Yrr < 0 Then kpe = kpen * 0.1
 End Function

 ''''''''''''''''''''''SOPR podprog.''''''''''''''''''''''''''''''''
Function podSOPR() As Single
'LEV
  If nnpl = 1 Then Xl = Xkp1: Yl = Ykp1: hl = hkp1
  If nnpl = 2 Then Xl = Xkp2: Yl = Ykp2: hl = hkp2
  If nnpl = 3 Then Xl = Xkp3: Yl = Ykp3: hl = hkp3
  If nnpl = 4 Then Xl = Xkp4: Yl = Ykp4: hl = hkp4
'PRAV
  If nnpp = 1 Then Xp = Xkp1: Yp = Ykp1: hp = hkp1
  If nnpp = 2 Then Xp = Xkp2: Yp = Ykp2: hp = hkp2
  If nnpp = 3 Then Xp = Xkp3: Yp = Ykp3: hp = hkp3
  If nnpp = 4 Then Xp = Xkp4: Yp = Ykp4: hp = hkp4
  dxso = Xp - Xl: dyso = Yp - Yl
  baz = Sqr(dxso ^ 2 + dyso ^ 2)
  aso = Abs(Atn(dyso / (dxso + 0.1)) / 3.141592 * 30) * 100

  If dxso > 0 And dyso > 0 Then Ygol.baz = Int(aso)
  If dxso < 0 And dyso > 0 Then Ygol.baz = Int(3000 - aso)
  If dxso < 0 And dyso < 0 Then Ygol.baz = Int(3000 + aso)
  If dxso > 0 And dyso < 0 Then Ygol.baz = Int(6000 - aso)
  If Alev < 1500 And Aprav > 4500 Then
  fi = Abs(Alev + 6000 - Aprav)
  ElseIf Alev > 4500 And Aprav < 1500 Then
  fi = Abs(Alev - (Aprav + 6000))
  Else
   fi = Abs(Alev - Aprav)
  End If
  If Alev < 1500 And Ygol.baz > 4500 Then
  blev = Abs(Alev + 6000 - Ygol.baz)
  ElseIf Alev > 4500 And Ygol.baz < 1500 Then
  blev = Abs(Alev - (Ygol.baz + 6000))
  Else
  blev = Abs(Alev - Ygol.baz)
  End If

  If Ygol.baz - 3000 < 0 Then
  ybazp = Ygol.baz + 3000
  Else
  ybazp = Ygol.baz - 3000
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
  Xc = Xcso: Yc = Ycso
  End Function

  ''''''''''''''''''''''dV0 podprog''''''''''''''''''''''''''''''''''''''''''''''
Function poddV0(ByVal tz As Single, ByVal zar As String, dv0) As Single
Dim t1 As Single, t2 As Single, t3 As Single, t4 As Single, t5 As Single, tz0 As Single
Dim tz11 As Single, dV01 As Single

Open App.Path & "\tz2c" For Input As #1
324: If EOF(1) Then GoTo 327
325: Input #1, t1, t2, t3, t4, t5
326: If t1 <= tz Or EOF(1) Then GoTo 32611
 GoTo 324
32611: If zar = "Полн" Then
        dv0 = t2
      ElseIf zar = "Умен" Or zar = "Перв" Then
      dv0 = t3
      ElseIf zar = "Втор" Or zar = "Трет" Then
      dv0 = t4
      ElseIf zar = "Четверт" Then
      dv0 = t5
      Else
      dv0 = t2
      End If
      tz0 = t1
327: Close #1
If tz < 0 Then tz11 = tz - 5
If tz >= 0 Then tz11 = tz + 5
    Open App.Path & "\tz2c" For Input As #1
424: If EOF(1) Then GoTo 427
425: Input #1, t1, t2, t3, t4, t5
426: If t1 <= tz11 Or EOF(1) Then GoTo 42611
 GoTo 424
42611: If zar = "Полн" Then
        dV01 = t2
      ElseIf zar = "Умен" Or zar = "Перв" Then
      dV01 = t3
      ElseIf zar = "Втор" Or zar = "Трет" Then
      dV01 = t4
      ElseIf zar = "Четверт" Then
      dV01 = t5
      Else
      dV01 = t2
      End If
427: Close #1
 If tz >= 0 Then dv0 = (dV01 - dv0) / 5 * (tz - tz0) + dv0
 If tz < 0 Then dv0 = ((dV01 - dv0) / 5 * (tz - tz0) - dv0) * -1
End Function
'''''''''''''''''''''''''''''''AR-5 podprog''''''''''''''''''''''''''''''''''
Function podAR5(ByVal zar As String, ByVal Disch As Single, Pricisch, N) As Single
If Disch = 0 Then Disch = 1
  If zar = "Полн" Then
  Open App.Path & "\ar5-P" For Input As #1
  ElseIf zar = "Умен" Then
  Open App.Path & "\ar5-Y" For Input As #1
  ElseIf zar = "Перв" Then
  Open App.Path & "\ar5-1" For Input As #1
  ElseIf zar = "Втор" Then
  Open App.Path & "\ar5-2" For Input As #1
  ElseIf zar = "Трет" Then
  Open App.Path & "\ar5-3" For Input As #1
  ElseIf zar = "Четверт" Then
  Open App.Path & "\ar5-4" For Input As #1
  Else
  Open App.Path & "\ar5-P" For Input As #1
  End If
301 If EOF(1) Then GoTo 303
  Input #1, ta1, ta2, ta3
  If ta1 <= Disch Or EOF(1) Then GoTo 302
      GoTo 301
302 Pric = ta2: Dta = ta1: N = ta3
303 Close #1
  Dischh = Disch + 200
  If zar = "Полн" Then
  Open App.Path & "\ar5-P" For Input As #1
  ElseIf zar = "Умен" Then
  Open App.Path & "\ar5-Y" For Input As #1
  ElseIf zar = "Перв" Then
  Open App.Path & "\ar5-1" For Input As #1
  ElseIf zar = "Втор" Then
  Open App.Path & "\ar5-2" For Input As #1
  ElseIf zar = "Трет" Then
  Open App.Path & "\ar5-3" For Input As #1
  ElseIf zar = "Четверт" Then
  Open App.Path & "\ar5-4" For Input As #1
  Else
  Open App.Path & "\ar5-P" For Input As #1
  End If
304 If EOF(1) Then GoTo 306
   Input #1, ta1, ta2, ta3
   If ta1 <= Dischh Or EOF(1) Then GoTo 305
   GoTo 304
305 Pricc = ta2: Nn = ta3
306  Close #1
    Pricisch = (Pricc - Pric) / 200 * (Disch - Dta) + Pric
    N = (Nn - N) / 200 * (Disch - Dta) + N
 End Function

'''''''''''''''''''''''''''3SH1 podprog'''''''''''''''''''''''''''''''''''''''''
Function pod3SH1(ByVal Disch As Single, ByVal zar As String, ByVal rep As String, ByVal vsem As Single, Pricisch, N, dNtus) As Single
If Disch > 15200 Then Disch = 15200
If zar = "Полн" Then
  Open App.Path & "\3sh-P" For Input As #1
  Else
   Open App.Path & "\3sh-Y" For Input As #1
End If
401 If EOF(1) Then GoTo 403
  Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8
  If ta1 <= Disch And ta1 + 200 >= Disch Or EOF(1) Then GoTo 402
      GoTo 401
402 Pric = ta2: Dta = ta1: N = ta3: dNtus = ta4: dNw = ta5: dNh = ta6: dNt = ta7: dNv0 = ta8
403 Close #1
  Dischh = Disch + 200
  If Dischh > 15000 Then Dischh = 15200
  If zar = "Полн" Then
  Open App.Path & "\3sh-P" For Input As #1
  Else
   Open App.Path & "\3sh-Y" For Input As #1
End If
404 If EOF(1) Then GoTo 406
   Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8
   If ta1 <= Dischh And ta1 + 200 >= Dischh Or EOF(1) Then GoTo 405
   GoTo 404
405 Pricc = ta2: Nn = ta3: dNww = ta5: dNhh = ta6: dNtt = ta7: dNv00 = ta8
406  Close #1
    Pricisch = (Pricc - Pric) / 200 * (Disch - Dta) + Pric
    dNw = (dNww - dNw) / 200 * (Disch - Dta) + dNw
    dNh = (dNhh - dNh) / 200 * (Disch - Dta) + dNh
    dNt = (dNtt - dNt) / 200 * (Disch - Dta) + dNt
    dNv0 = (dNv00 - dNvo) / 200 * (Disch - Dta) + dNv0
    If rep = "Пристрелян" Then
        dN = dN
        Else
        dN = ((dNw / 10) * Wx) + ((dNh / 10 * -1) * dhh) + ((dNt / 10) * dddt) + (dNv0 * dV00)
    End If
    N = (Nn - N) / 200 * (Disch - Dta) + N + dN - 3
 End Function
'''''''''''''''''''''''''''B-90 podprog''''''''''''''''''''''''''''''''''''''''''''
Function podB90(ByVal zar As String, ByVal Disch As Single, ByVal rep As String, ByVal Wx As Single, N, dNtus, vrv, Pricisch) As Single
    If zar = "Полн" Then
    Open App.Path & "\B-90p" For Input As #1
    ElseIf zar = "Умен" Then
    Open App.Path & "\B-90Y" For Input As #1
    ElseIf zar = "Перв" Then
    Open App.Path & "\B-901" For Input As #1
    ElseIf zar = "Втор" Then
    Open App.Path & "\B-902" For Input As #1
    Else
    Open App.Path & "\B-903" For Input As #1
    End If
501 If EOF(1) Then GoTo 503
  Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8
  If Disch > 14800 Then Disch = 14800
  If ta1 <= Disch And ta1 + 200 >= Disch Or EOF(1) Then GoTo 502
      GoTo 501
502 Pric = ta2: Dta = ta1: N = ta3: dNtus = ta4: dNw = ta5: dNh = ta6: dNt = ta7: dNv0 = ta8
503 Close #1
  Dischh = Disch + 200
  If Dischh > 14800 Then Dischh = 14800
    If zar = "Полн" Then
    Open App.Path & "\B-90p" For Input As #1
    ElseIf zar = "Умен" Then
    Open App.Path & "\B-90Y" For Input As #1
    ElseIf zar = "Перв" Then
    Open App.Path & "\B-901" For Input As #1
    ElseIf zar = "Втор" Then
    Open App.Path & "\B-902" For Input As #1
    Else
    Open App.Path & "\B-903" For Input As #1
    End If
504 If EOF(1) Then GoTo 506
   Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8
   If ta1 <= Dischh And ta1 + 200 >= Dischh Or EOF(1) Then GoTo 505
   GoTo 504
505 Pricc = ta2: Nn = ta3: dNww = ta5: dNhh = ta6: dNtt = ta7: dNv00 = ta8
506  Close #1
    Pricisch = (Pricc - Pric) / 200 * (Disch - Dta) + Pric + 2
    dNw = (dNww - dNw) / 200 * (Disch - Dta) + dNw
    dNh = (dNhh - dNh) / 200 * (Disch - Dta) + dNh
    dNt = (dNtt - dNt) / 200 * (Disch - Dta) + dNt
    dNv0 = (dNv00 - dNv0) / 200 * (Disch - Dta) + dNv0
    If rep = "Пристрелян" Then
        dN = dN
        Else
        dN = ((dNw / 10) * Wx) + ((dNh / 10 * -1) * dhh) + ((dNt / 10) * dddt) + (dNv0 * dV00)
    End If
    N = (Nn - N) / 200 * (Disch - Dta) + N + dN
510 If zar = "Полн" And Disch >= 5200 Then
  vrv = 15
 ElseIf zar = "Умен" And Disch >= 4400 Then
         vrv = 15
                ElseIf zar = "Перв" And Disch >= 3800 Then
                vrv = 15
                        ElseIf zar = "Втор" And Disch >= 3400 Then
                       
                        vrv = 15
                                ElseIf zar = "Трет" And Disch >= 3000 Then
                                vrv = 15
         Else
          vrv = 0
 End If
End Function
''''''''''''''''''''''''''''''T-90 podprogr''''''''''''''''''''''''''''''''''
Function podT90(ByVal Disch As Single, ByVal zar As String, N, dNtus, Pricisch) As Single
If Disch > 15000 Then Disch = 15000
    If zar = "Полн" Then
    Open App.Path & "\T-90p" For Input As #1
    ElseIf zar = "Умен" Then
    Open App.Path & "\T-90Y" For Input As #1
    ElseIf zar = "Перв" Then
    Open App.Path & "\T-901" For Input As #1
    ElseIf zar = "Втор" Then
    Open App.Path & "\T-902" For Input As #1
    ElseIf zar = "Трет" Then
    Open App.Path & "\T-903" For Input As #1
    Else
    Open App.Path & "\T-904" For Input As #1
    End If
601 If EOF(1) Then GoTo 603
  Input #1, ta1, ta2, ta3, ta4
  If ta1 <= Disch And ta1 + 200 >= Disch Or EOF(1) Then GoTo 602
      GoTo 601
602 Pric = ta2: Dta = ta1: N = ta3: dNtus = ta4
603 Close #1
  Dischh = Disch + 200
  If Dischh > 15000 Then Dischh = 15000
    If zar = "Полн" Then
    Open App.Path & "\T-90p" For Input As #1
    ElseIf zar = "Умен" Then
    Open App.Path & "\T-90Y" For Input As #1
    ElseIf zar = "Перв" Then
    Open App.Path & "\T-901" For Input As #1
    ElseIf zar = "Втор" Then
    Open App.Path & "\T-902" For Input As #1
    ElseIf zar = "Трет" Then
    Open App.Path & "\T-903" For Input As #1
    Else
    Open App.Path & "\T-904" For Input As #1
    End If
604 If EOF(1) Then GoTo 606
   Input #1, ta1, ta2, ta3, ta4
   If ta1 <= Dischh And ta1 + 200 >= Dischh Or EOF(1) Then GoTo 605
   GoTo 604
605 Pricc = ta2: Nn = ta3
606  Close #1
    Pricisch = (Pricc - Pric) / 200 * (Disch - Dta) + Pric
    N = (Nn - N) / 200 * (Disch - Dta) + N
 End Function

''''''''''''''''''''''''MORTIRNAIA''''''''''''''''''''''''''''''''''''''''''''''
Function podPOPRMORT(ByVal zar As String, ByVal Dt As Single, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, ByVal met As Single, ybylc, Wx, Wz, dddt) As Single
   If zar = "Полн" Then
   Open App.Path & "\2C1MORTP" For Input As #1
   ElseIf zar = "Умен" Then
   Open App.Path & "\2C1MORTY" For Input As #1
   ElseIf zar = "Перв" Then
   Open App.Path & "\2C1MORT1" For Input As #1
   ElseIf zar = "Втор" Then
   Open App.Path & "\2C1MORT2" For Input As #1
   ElseIf zar = "Трет" Then
   Open App.Path & "\2C1MORT3" For Input As #1
   ElseIf zar = "Четверт" Then
   Open App.Path & "\2C1MORT4" For Input As #1
   Else
   Open App.Path & "\2C1MORTP" For Input As #1
   End If
391 If EOF(1) Then GoTo 39212
39111 Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8, ta9, ta10
392 If ta1 <= Dt Or EOF(1) Then GoTo 39211
 GoTo 391
39211 Dta = ta1: Prta = ta2: dXtus = ta3: z = ta4: dzw = ta5: dxw = ta6
 dxh = ta7: dxt = ta8: dxv0 = ta9: ybyl = 40: ts = ta10

39212 Close #1
   dtT = Dt + 200
   If zar = "Полн" Then
  Open App.Path & "\2C1MORTP" For Input As #1
   ElseIf zar = "Умен" Then
   Open App.Path & "\2C1MORTY" For Input As #1
   ElseIf zar = "Перв" Then
   Open App.Path & "\2C1MORT1" For Input As #1
   ElseIf zar = "Втор" Then
   Open App.Path & "\2C1MORT2" For Input As #1
   ElseIf zar = "Трет" Then
   Open App.Path & "\2C1MORT3" For Input As #1
   ElseIf zar = "Четверт" Then
   Open App.Path & "\2C1MORT4" For Input As #1
   Else
   Open App.Path & "\2C1MORTP" For Input As #1
   End If
3911 If EOF(1) Then GoTo 392121
391111 Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8, ta9, ta10

3921 If ta1 <= dtT Or EOF(1) Then GoTo 392111
 GoTo 3911
392111 dta1 = ta1: prta1 = ta2: dXtus1 = ta3: z1 = ta4: dzw1 = ta5: dxw1 = ta6
 dxh1 = ta7: dxt1 = ta8: dxv01 = ta9

392121 Close #1

zc = ((z1 - z) / 200 * (Dt - Dta) + z) * -1
dZwc = ((dzw1 - dzw) / 200 * (Dt - Dta) + dzw) * -1 / 10
dXwc = ((dxw1 - dxw) / 200 * (Dt - Dta) + dxw) * -1 / 10
dXhc = ((dxw1 - dxw) / 200 * (Dt - Dta) + dxh) / 10
dXtc = ((dxt1 - dxt) / 200 * (Dt - Dta) + dxt) * -1 / 10
dXv0c = ((dxv01 - dxv0) / 200 * (Dt - Dta) + dxv0) * -1
If met = 0 Then
 Aw200 = BP.pAw02: Aw400 = BP.pAw04: Aw800 = BP.pAw08: Aw1200 = BP.pAw12: Aw1600 = BP.pAw16
 Aw2000 = BP.pAw20: Aw2400 = BP.pAw24: Aw3000 = BP.pAw30: Aw4000 = BP.pAw40
Else
End If
  ybylc = 4000
 If ybylc = 4000 Then aww = Aw4000 * 100
 Ygolt = OZ.Ygolt1
If Ygolt - aww < 0 Then
       aws = Ygolt + 6000 - aww
   Else
   aws = Ygolt - aww
   End If
  ybylc = 4000
 If ybylc = 4000 Then sw = BP.pW40

 Wx = Abs(sw * Cos((aws / 100 * 6) * 3.141596 / 180))
 Wz = Abs(sw * Sin((aws / 100 * 6) * 3.14159 / 180))

 If aws <= 1500 Or aws >= 4500 Then Wx = Wx * -1
 If aws >= 3000 Then Wz = Wz * -1

  ybylc = 4000
 If ybylc = 4000 Then dddt = BP.pdT40
 If dddt >= 50 Then dddt = (dddt - 50) * -1
End Function
'''''''''''''''''''''PRICEL MORTIRNAIA''''''''''''''''''''''''''''''''''''''''
Function podPRICMORTRGM(ByVal zar As String, ByVal Disch As Single, Pricisch, ts) As Single
  If zar = "Полн" Then
  Open App.Path & "\2C1MORTP" For Input As #1
   ElseIf zar = "Умен" Then
   Open App.Path & "\2C1MORTY" For Input As #1
   ElseIf zar = "Перв" Then
   Open App.Path & "\2C1MORT1" For Input As #1
   ElseIf zar = "Втор" Then
   Open App.Path & "\2C1MORT2" For Input As #1
   ElseIf zar = "Трет" Then
   Open App.Path & "\2C1MORT3" For Input As #1
   ElseIf zar = "Четверт" Then
   Open App.Path & "\2C1MORT4" For Input As #1
   Else
   Open App.Path & "\2C1MORTP" For Input As #1
   End If
392131 If EOF(1) Then GoTo 392132
  Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8, ta9, ta10
  If ta1 <= Disch Or EOF(1) Then GoTo 3921311
      GoTo 392131
3921311 Pric = ta2: Dta = ta1: ts = ta10
392132 Close #1
  Dischh = Disch + 200
 If zar = "Полн" Then
  Open App.Path & "\2C1MORTP" For Input As #1
   ElseIf zar = "Умен" Then
   Open App.Path & "\2C1MORTY" For Input As #1
   ElseIf zar = "Перв" Then
   Open App.Path & "\2C1MORT1" For Input As #1
   ElseIf zar = "Втор" Then
   Open App.Path & "\2C1MORT2" For Input As #1
   ElseIf zar = "Трет" Then
   Open App.Path & "\2C1MORT3" For Input As #1
   ElseIf zar = "Четверт" Then
   Open App.Path & "\2C1MORT4" For Input As #1
   Else
   Open App.Path & "\2C1MORTP" For Input As #1
   End If
3921321 If EOF(1) Then GoTo 3921322
   Input #1, ta1, ta2, ta3, ta4, ta5, ta6, ta7, ta8, ta9, ta10
   If ta1 <= Dischh Or EOF(1) Then GoTo 39213211
   GoTo 3921321
39213211 Pricc = ta2
3921322  Close #1
    Pricisch = (Pricc - Pric) / 200 * (Disch - Dta) + Pric
End Function
''''''''''''''''''''''''''''''KPE'''''''''''''''''''''''''''''''''''''''''
Function podKPEmort(ByVal zar As String, ByVal Pricisch As Single, ByVal Yrr As Single, kpe) As Single
If zar = "Полн" Then
        Open App.Path & "\2c1kpemp" For Input As #1
        ElseIf zar = "Умен" Then
        Open App.Path & "\2c1kpemy" For Input As #1
        ElseIf zar = "Перв" Then
        Open App.Path & "\2c1kpem1" For Input As #1
        ElseIf zar = "Втор" Then
        Open App.Path & "\2c1kpem2" For Input As #1
        ElseIf zar = "Трет" Then
        Open App.Path & "\2c1kpem3" For Input As #1
        ElseIf zar = "Четверт" Then
        Open App.Path & "\2c1kpem4" For Input As #1
        Else
        Open App.Path & "\2c1kpemp" For Input As #1
End If
3921332 If EOF(1) Then GoTo 3921333
  Input #1, e1, e2, e3
   If Pricisch >= e1 And e1 + 20 >= Pricisch Then kpew = e2: kpen = e3: GoTo 3921333
   GoTo 3921332
3921333 Close #1
    If Yrr >= 0 Then kpe = kpew * 0.1
    If Yrr < 0 Then kpe = kpen * 0.1
  End Function
  
Function YGLUKNP(Dtkp, Yrkp, Ygoltkp, ByVal Xkp As Single, ByVal Ykp As Single, ByVal hkp As Single) As Single
  Dim Xc As Single, Yc As Single, hc As Single
  Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc
   dx = Xc - Xkp
 dy = Yc - Ykp
 dh = hc - hkp
   Pi = 3.14159265358
 Dtkp = Round(Sqr(dx ^ 2 + dy ^ 2))
 Yrkp = Round((dh / (Dtkp * 0.001 + 0.1)) * 0.95)
 a = Abs(Atn(dy / (dx + 0.1)) / Pi * 30) * 100
 If dx > 0 And dy > 0 Then Ygoltkp = Int(a)
 If dx < 0 And dy > 0 Then Ygoltkp = Int(3000 - a)
 If dx < 0 And dy < 0 Then Ygoltkp = Int(3000 + a)
 If dx > 0 And dy < 0 Then Ygoltkp = Int(6000 - a)
End Function

Private Sub phc_Change()
Shest6Oryd.phc = phc
End Sub

Private Sub pplZel_Click()
Dim z(1 To 10) As String
Dim nz As String
Dim Xc As Single, Yc As Single, hc As Single, Fr As Single, Gl As Single
nz = pplZel
1011 Open "D:\YO_NA\Zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, z(1), z(2), z(3), z(4), z(5), z(6)
   If z(1) = nz Then Xc = z(2): Yc = z(3): hc = Val(z(4)): Fr = z(5): Gl = z(6): GoTo 1012
        GoTo 101111
1012 Close #1
pXc.Text = Xc: pYc.Text = Yc: phc.Text = hc: pFrontc.Text = Fr: pGlybinac.Text = Gl
End Sub
Private Sub pplZel_KeyDown(KeyCode As Integer, Shift As Integer)
Dim z(1 To 10) As String
Dim nz As String
Dim Xc As Single, Yc As Single, hc As Single
nz = pplZel
If KeyCode = 13 Then
1011    Open "D:\YO_NA\zeli" For Input As #1
101111  If EOF(1) Then GoTo 1012
   Input #1, z(1), z(2), z(3), z(4), z(5), z(6)
   If z(1) = nz Then Xc = z(2): Yc = z(3): hc = z(4): Fr = z(5): Gl = z(6): GoTo 1012
        GoTo 101111
1012    Close #1
pXc.Text = Xc: pYc.Text = Yc: phc.Text = hc: pFrontc.Text = Fr: pGlybinac.Text = Gl
    Else
End If
End Sub

Private Sub pXc_Change()
Shest6Oryd.pXc = pXc
End Sub
Private Sub pYc_Change()
Shest6Oryd.pYc = pYc
End Sub

Private Sub ZasGrZ_Click()
zasGryp.Show
End Sub
Private Sub pXc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc.Text = ""
pYc.SetFocus
End If
End Sub
Private Sub pyc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc.Text = ""
phc.SetFocus
End If
End Sub
Private Sub pFrontc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pGlybinac.Text = ""
pGlybinac.SetFocus
End If
End Sub
Private Sub pAc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pDc.Text = ""
pDc.SetFocus
End If
End Sub
Private Sub pDc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMc.Text = ""
pMc.SetFocus
End If
End Sub
Private Sub pAlc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMlc.Text = ""
pMlc.SetFocus
End If
End Sub
Private Sub pApc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMpc.Text = ""
pMpc.SetFocus
End If
End Sub

Sub msgVelikaDalnost(ByVal snar As String, ByVal zar As String, ByVal Nop As String, ByVal Dt1 As Double)
'сообщение велика дальность
Dim mesage As String, mesage2 As String, b As String

mesage = "Для " & Nop & " и выбранного заряда велика дальность"

If snar = "3ОФ56" Then
        If zar = "Полн" And Dt1 > 16790 Then b = MsgBox(mesage, vbOKOnly, "ВЕЛИКА ДАЛЬНОСТЬ")
    Else
        If zar = "Полн" And Dt1 > 15200 Then b = MsgBox(mesage, vbOKOnly, "ВЕЛИКА ДАЛЬНОСТЬ")
        If zar = "Умен" And Dt1 > 12800 Then b = MsgBox(mesage, vbOKOnly, "ВЕЛИКА ДАЛЬНОСТЬ")
        If zar = "Перв" And Dt1 > 11800 Then b = MsgBox(mesage, vbOKOnly, "ВЕЛИКА ДАЛЬНОСТЬ")
        If zar = "Втор" And Dt1 > 10000 Then b = MsgBox(mesage, vbOKOnly, "ВЕЛИКА ДАЛЬНОСТЬ")
        If zar = "Трет" And Dt1 > 8600 Then b = MsgBox(mesage, vbOKOnly, "ВЕЛИКА ДАЛЬНОСТЬ")
        If zar = "Четверт" And Dt1 > 6400 Then b = MsgBox(mesage, vbOKOnly, "ВЕЛИКА ДАЛЬНОСТЬ")
End If

End Sub
