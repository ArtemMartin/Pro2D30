VERSION 5.00
Begin VB.Form BP 
   BackColor       =   &H0000C0C0&
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Times New Roman"
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
   Begin VB.Frame Frame4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Выбор Снаряд. Взрыватель. Заряд."
      Height          =   2800
      Left            =   7080
      TabIndex        =   156
      Top             =   6900
      Width           =   7400
      Begin VB.ComboBox Combo12 
         Height          =   405
         ItemData        =   "Form1.frx":0000
         Left            =   5500
         List            =   "Form1.frx":0016
         TabIndex        =   175
         Text            =   "Полн"
         Top             =   1900
         Width           =   1200
      End
      Begin VB.ComboBox Combo11 
         Height          =   405
         ItemData        =   "Form1.frx":003E
         Left            =   4100
         List            =   "Form1.frx":0054
         TabIndex        =   174
         Text            =   "Полн"
         Top             =   1900
         Width           =   1200
      End
      Begin VB.ComboBox Combo10 
         Height          =   405
         ItemData        =   "Form1.frx":007C
         Left            =   2700
         List            =   "Form1.frx":0092
         TabIndex        =   173
         Text            =   "Полн"
         Top             =   1900
         Width           =   1200
      End
      Begin VB.ComboBox Combo9 
         Height          =   405
         ItemData        =   "Form1.frx":00BA
         Left            =   1300
         List            =   "Form1.frx":00D0
         TabIndex        =   172
         Text            =   "Полн"
         Top             =   1900
         Width           =   1200
      End
      Begin VB.ComboBox Combo8 
         Height          =   405
         ItemData        =   "Form1.frx":00F8
         Left            =   5500
         List            =   "Form1.frx":010B
         TabIndex        =   171
         Text            =   "РГМ"
         Top             =   1400
         Width           =   1200
      End
      Begin VB.ComboBox Combo7 
         Height          =   405
         ItemData        =   "Form1.frx":012A
         Left            =   4100
         List            =   "Form1.frx":013D
         TabIndex        =   170
         Text            =   "РГМ"
         Top             =   1400
         Width           =   1200
      End
      Begin VB.ComboBox Combo6 
         Height          =   405
         ItemData        =   "Form1.frx":015C
         Left            =   2700
         List            =   "Form1.frx":016F
         TabIndex        =   169
         Text            =   "РГМ"
         Top             =   1400
         Width           =   1200
      End
      Begin VB.ComboBox Combo5 
         Height          =   405
         ItemData        =   "Form1.frx":018E
         Left            =   1300
         List            =   "Form1.frx":01A1
         TabIndex        =   168
         Text            =   "РГМ"
         Top             =   1400
         Width           =   1200
      End
      Begin VB.ComboBox Combo4 
         Height          =   405
         ItemData        =   "Form1.frx":01C0
         Left            =   5500
         List            =   "Form1.frx":01D0
         TabIndex        =   167
         Text            =   "ОФ"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox Combo3 
         Height          =   405
         ItemData        =   "Form1.frx":01E4
         Left            =   4100
         List            =   "Form1.frx":01F4
         TabIndex        =   166
         Text            =   "ОФ"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox Combo2 
         Height          =   405
         ItemData        =   "Form1.frx":0208
         Left            =   2700
         List            =   "Form1.frx":0218
         TabIndex        =   165
         Text            =   "ОФ"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         ItemData        =   "Form1.frx":022C
         Left            =   1300
         List            =   "Form1.frx":023C
         TabIndex        =   164
         Text            =   "ОФ"
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label Label67 
         BackColor       =   &H0000C0C0&
         Caption         =   "Заряд"
         Height          =   300
         Left            =   300
         TabIndex        =   163
         Top             =   1900
         Width           =   800
      End
      Begin VB.Label Label66 
         BackColor       =   &H0000C0C0&
         Caption         =   "Взрыв."
         Height          =   300
         Left            =   300
         TabIndex        =   162
         Top             =   1400
         Width           =   800
      End
      Begin VB.Label Label65 
         BackColor       =   &H0000C0C0&
         Caption         =   "Снаряд"
         Height          =   300
         Left            =   300
         TabIndex        =   161
         Top             =   900
         Width           =   800
      End
      Begin VB.Label Label64 
         BackColor       =   &H0000C0C0&
         Caption         =   "4 Бат"
         Height          =   300
         Left            =   5500
         TabIndex        =   160
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label63 
         BackColor       =   &H0000C0C0&
         Caption         =   "3 Бат"
         Height          =   300
         Left            =   4100
         TabIndex        =   159
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label62 
         BackColor       =   &H0000C0C0&
         Caption         =   "2 Бат"
         Height          =   300
         Left            =   2700
         TabIndex        =   158
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label61 
         BackColor       =   &H0000C0C0&
         Caption         =   "1 Бат"
         Height          =   300
         Left            =   1300
         TabIndex        =   157
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000C0C0&
      Caption         =   " Метеоданные"
      Height          =   6600
      Left            =   12600
      TabIndex        =   98
      Top             =   50
      Width           =   5700
      Begin VB.TextBox Text93 
         Height          =   400
         Left            =   4600
         TabIndex        =   155
         Text            =   "0"
         Top             =   5800
         Width           =   600
      End
      Begin VB.TextBox Text92 
         Height          =   400
         Left            =   4600
         TabIndex        =   154
         Text            =   "0"
         Top             =   5300
         Width           =   600
      End
      Begin VB.TextBox Text91 
         Height          =   400
         Left            =   4600
         TabIndex        =   153
         Text            =   "0"
         Top             =   4800
         Width           =   600
      End
      Begin VB.TextBox Text90 
         Height          =   400
         Left            =   3800
         TabIndex        =   152
         Text            =   "0"
         Top             =   5800
         Width           =   600
      End
      Begin VB.TextBox Text89 
         Height          =   400
         Left            =   3800
         TabIndex        =   151
         Text            =   "0"
         Top             =   5300
         Width           =   600
      End
      Begin VB.TextBox Text88 
         Height          =   400
         Left            =   3800
         TabIndex        =   150
         Text            =   "0"
         Top             =   4800
         Width           =   600
      End
      Begin VB.TextBox Text87 
         Height          =   400
         Left            =   3000
         TabIndex        =   149
         Text            =   "0"
         Top             =   5800
         Width           =   600
      End
      Begin VB.TextBox Text86 
         Height          =   400
         Left            =   3000
         TabIndex        =   148
         Text            =   "0"
         Top             =   5300
         Width           =   600
      End
      Begin VB.TextBox Text85 
         Height          =   400
         Left            =   3000
         TabIndex        =   147
         Text            =   "0"
         Top             =   4800
         Width           =   600
      End
      Begin VB.TextBox Text84 
         Height          =   400
         Left            =   4600
         TabIndex        =   146
         Text            =   "0"
         Top             =   4300
         Width           =   600
      End
      Begin VB.TextBox Text83 
         Height          =   400
         Left            =   4600
         TabIndex        =   145
         Text            =   "0"
         Top             =   3800
         Width           =   600
      End
      Begin VB.TextBox Text82 
         Height          =   400
         Left            =   4600
         TabIndex        =   144
         Text            =   "0"
         Top             =   3300
         Width           =   600
      End
      Begin VB.TextBox Text81 
         Height          =   400
         Left            =   4600
         TabIndex        =   143
         Text            =   "0"
         Top             =   2800
         Width           =   600
      End
      Begin VB.TextBox Text80 
         Height          =   400
         Left            =   4600
         TabIndex        =   142
         Text            =   "0"
         Top             =   2300
         Width           =   600
      End
      Begin VB.TextBox Text79 
         Height          =   400
         Left            =   4600
         TabIndex        =   141
         Text            =   "0"
         Top             =   1800
         Width           =   600
      End
      Begin VB.TextBox Text78 
         Height          =   400
         Left            =   3800
         TabIndex        =   140
         Text            =   "0"
         Top             =   4300
         Width           =   600
      End
      Begin VB.TextBox Text77 
         Height          =   400
         Left            =   3800
         TabIndex        =   139
         Text            =   "0"
         Top             =   3800
         Width           =   600
      End
      Begin VB.TextBox Text76 
         Height          =   400
         Left            =   3800
         TabIndex        =   138
         Text            =   "0"
         Top             =   3300
         Width           =   600
      End
      Begin VB.TextBox Text75 
         Height          =   400
         Left            =   3800
         TabIndex        =   137
         Text            =   "0"
         Top             =   2800
         Width           =   600
      End
      Begin VB.TextBox Text74 
         Height          =   400
         Left            =   3800
         TabIndex        =   136
         Text            =   "0"
         Top             =   2300
         Width           =   600
      End
      Begin VB.TextBox Text73 
         Height          =   400
         Left            =   3800
         TabIndex        =   135
         Text            =   "0"
         Top             =   1800
         Width           =   600
      End
      Begin VB.TextBox Text72 
         Height          =   400
         Left            =   3000
         TabIndex        =   134
         Text            =   "0"
         Top             =   4300
         Width           =   600
      End
      Begin VB.TextBox Text71 
         Height          =   400
         Left            =   3000
         TabIndex        =   133
         Text            =   "0"
         Top             =   3800
         Width           =   600
      End
      Begin VB.TextBox Text70 
         Height          =   400
         Left            =   3000
         TabIndex        =   132
         Text            =   "0"
         Top             =   3300
         Width           =   600
      End
      Begin VB.TextBox Text69 
         Height          =   400
         Left            =   3000
         TabIndex        =   131
         Text            =   "0"
         Top             =   2800
         Width           =   600
      End
      Begin VB.TextBox Text68 
         Height          =   400
         Left            =   3000
         TabIndex        =   130
         Text            =   "0"
         Top             =   2300
         Width           =   600
      End
      Begin VB.TextBox Text67 
         Height          =   400
         Left            =   3000
         TabIndex        =   129
         Text            =   "0"
         Top             =   1800
         Width           =   600
      End
      Begin VB.TextBox Text66 
         Height          =   400
         Left            =   4600
         TabIndex        =   128
         Text            =   "0"
         Top             =   1300
         Width           =   600
      End
      Begin VB.TextBox Text65 
         Height          =   400
         Left            =   3800
         TabIndex        =   127
         Text            =   "0"
         Top             =   1300
         Width           =   600
      End
      Begin VB.TextBox Text64 
         Height          =   400
         Left            =   3000
         TabIndex        =   126
         Text            =   "0"
         Top             =   1300
         Width           =   600
      End
      Begin VB.TextBox Text63 
         Height          =   400
         Left            =   4600
         TabIndex        =   125
         Text            =   "0"
         Top             =   800
         Width           =   600
      End
      Begin VB.TextBox Text62 
         Height          =   400
         Left            =   3800
         TabIndex        =   124
         Text            =   "0"
         Top             =   800
         Width           =   600
      End
      Begin VB.TextBox Text61 
         Height          =   400
         Left            =   3000
         TabIndex        =   123
         Text            =   "0"
         Top             =   800
         Width           =   600
      End
      Begin VB.TextBox Text60 
         Height          =   400
         Left            =   1300
         TabIndex        =   108
         Text            =   "0"
         Top             =   2500
         Width           =   1000
      End
      Begin VB.TextBox Text59 
         Height          =   400
         Left            =   1300
         TabIndex        =   107
         Text            =   "0"
         Top             =   2000
         Width           =   1000
      End
      Begin VB.TextBox Text58 
         Height          =   400
         Left            =   1300
         TabIndex        =   106
         Text            =   "0"
         Top             =   1500
         Width           =   1000
      End
      Begin VB.TextBox Text57 
         Height          =   400
         Left            =   1300
         TabIndex        =   105
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox Text56 
         Height          =   400
         Left            =   1300
         TabIndex        =   104
         Text            =   "0"
         Top             =   500
         Width           =   1000
      End
      Begin VB.Label Label60 
         BackColor       =   &H0000C0C0&
         Caption         =   "60"
         Height          =   300
         Left            =   2600
         TabIndex        =   122
         Top             =   5800
         Width           =   300
      End
      Begin VB.Label Label59 
         BackColor       =   &H0000C0C0&
         Caption         =   "50"
         Height          =   300
         Left            =   2600
         TabIndex        =   121
         Top             =   5300
         Width           =   300
      End
      Begin VB.Label Label58 
         BackColor       =   &H0000C0C0&
         Caption         =   "40"
         Height          =   300
         Left            =   2600
         TabIndex        =   120
         Top             =   4800
         Width           =   300
      End
      Begin VB.Label Label57 
         BackColor       =   &H0000C0C0&
         Caption         =   "30"
         Height          =   300
         Left            =   2600
         TabIndex        =   119
         Top             =   4300
         Width           =   300
      End
      Begin VB.Label Label56 
         BackColor       =   &H0000C0C0&
         Caption         =   "24"
         Height          =   300
         Left            =   2600
         TabIndex        =   118
         Top             =   3800
         Width           =   300
      End
      Begin VB.Label Label55 
         BackColor       =   &H0000C0C0&
         Caption         =   "20"
         Height          =   300
         Left            =   2600
         TabIndex        =   117
         Top             =   3300
         Width           =   300
      End
      Begin VB.Label Label54 
         BackColor       =   &H0000C0C0&
         Caption         =   "16"
         Height          =   300
         Left            =   2600
         TabIndex        =   116
         Top             =   2800
         Width           =   300
      End
      Begin VB.Label Label53 
         BackColor       =   &H0000C0C0&
         Caption         =   "12"
         Height          =   300
         Left            =   2600
         TabIndex        =   115
         Top             =   2300
         Width           =   300
      End
      Begin VB.Label Label52 
         BackColor       =   &H0000C0C0&
         Caption         =   "08"
         Height          =   300
         Left            =   2600
         TabIndex        =   114
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label Label51 
         BackColor       =   &H0000C0C0&
         Caption         =   "04"
         Height          =   300
         Left            =   2600
         TabIndex        =   113
         Top             =   1300
         Width           =   300
      End
      Begin VB.Label Label50 
         BackColor       =   &H0000C0C0&
         Caption         =   "02"
         Height          =   300
         Left            =   2600
         TabIndex        =   112
         Top             =   800
         Width           =   300
      End
      Begin VB.Label Label49 
         BackColor       =   &H0000C0C0&
         Caption         =   "W"
         Height          =   300
         Left            =   4600
         TabIndex        =   111
         Top             =   360
         Width           =   400
      End
      Begin VB.Label Label48 
         BackColor       =   &H0000C0C0&
         Caption         =   "Aw"
         Height          =   300
         Left            =   3800
         TabIndex        =   110
         Top             =   360
         Width           =   400
      End
      Begin VB.Label Label47 
         BackColor       =   &H0000C0C0&
         Caption         =   "dT"
         Height          =   300
         Left            =   3000
         TabIndex        =   109
         Top             =   360
         Width           =   400
      End
      Begin VB.Label Label46 
         BackColor       =   &H0000C0C0&
         Caption         =   "W="
         Height          =   300
         Left            =   700
         TabIndex        =   103
         Top             =   2500
         Width           =   400
      End
      Begin VB.Label Label45 
         BackColor       =   &H0000C0C0&
         Caption         =   "Aw="
         Height          =   300
         Left            =   700
         TabIndex        =   102
         Top             =   2000
         Width           =   400
      End
      Begin VB.Label Label44 
         BackColor       =   &H0000C0C0&
         Caption         =   "Tз="
         Height          =   300
         Left            =   750
         TabIndex        =   101
         Top             =   1500
         Width           =   400
      End
      Begin VB.Label Label43 
         BackColor       =   &H0000C0C0&
         Caption         =   "H="
         Height          =   300
         Left            =   850
         TabIndex        =   100
         Top             =   1000
         Width           =   300
      End
      Begin VB.Label Label42 
         BackColor       =   &H0000C0C0&
         Caption         =   "h метео="
         Height          =   300
         Left            =   200
         TabIndex        =   99
         Top             =   500
         Width           =   1100
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Балистика"
      Height          =   6100
      Left            =   7050
      TabIndex        =   57
      Top             =   50
      Width           =   5300
      Begin VB.TextBox Text55 
         Height          =   400
         Left            =   4100
         TabIndex        =   97
         Text            =   "0"
         Top             =   5300
         Width           =   600
      End
      Begin VB.TextBox Text54 
         Height          =   400
         Left            =   3100
         TabIndex        =   96
         Text            =   "0"
         Top             =   5300
         Width           =   600
      End
      Begin VB.TextBox Text53 
         Height          =   400
         Left            =   2100
         TabIndex        =   95
         Text            =   "0"
         Top             =   5300
         Width           =   600
      End
      Begin VB.TextBox Text52 
         Height          =   400
         Left            =   1100
         TabIndex        =   94
         Text            =   "0"
         Top             =   5300
         Width           =   600
      End
      Begin VB.TextBox Text51 
         Height          =   400
         Left            =   4100
         TabIndex        =   93
         Text            =   "0"
         Top             =   4700
         Width           =   600
      End
      Begin VB.TextBox Text50 
         Height          =   400
         Left            =   3100
         TabIndex        =   92
         Text            =   "0"
         Top             =   4700
         Width           =   600
      End
      Begin VB.TextBox Text49 
         Height          =   400
         Left            =   2100
         TabIndex        =   91
         Text            =   "0"
         Top             =   4700
         Width           =   600
      End
      Begin VB.TextBox Text48 
         Height          =   400
         Left            =   1100
         TabIndex        =   90
         Text            =   "0"
         Top             =   4700
         Width           =   600
      End
      Begin VB.TextBox Text47 
         Height          =   400
         Left            =   4100
         TabIndex        =   89
         Text            =   "0"
         Top             =   4100
         Width           =   600
      End
      Begin VB.TextBox Text46 
         Height          =   400
         Left            =   3100
         TabIndex        =   88
         Text            =   "0"
         Top             =   4100
         Width           =   600
      End
      Begin VB.TextBox Text45 
         Height          =   400
         Left            =   2100
         TabIndex        =   87
         Text            =   "0"
         Top             =   4100
         Width           =   600
      End
      Begin VB.TextBox Text44 
         Height          =   400
         Left            =   1100
         TabIndex        =   86
         Text            =   "0"
         Top             =   4100
         Width           =   600
      End
      Begin VB.TextBox Text43 
         Height          =   400
         Left            =   4100
         TabIndex        =   85
         Text            =   "0"
         Top             =   3500
         Width           =   600
      End
      Begin VB.TextBox Text42 
         Height          =   400
         Left            =   3100
         TabIndex        =   84
         Text            =   "0"
         Top             =   3500
         Width           =   600
      End
      Begin VB.TextBox Text41 
         Height          =   400
         Left            =   2100
         TabIndex        =   83
         Text            =   "0"
         Top             =   3500
         Width           =   600
      End
      Begin VB.TextBox Text40 
         Height          =   400
         Left            =   1100
         TabIndex        =   82
         Text            =   "0"
         Top             =   3500
         Width           =   600
      End
      Begin VB.TextBox Text39 
         Height          =   400
         Left            =   4100
         TabIndex        =   81
         Text            =   "0"
         Top             =   2900
         Width           =   600
      End
      Begin VB.TextBox Text38 
         Height          =   400
         Left            =   3100
         TabIndex        =   80
         Text            =   "0"
         Top             =   2900
         Width           =   600
      End
      Begin VB.TextBox Text37 
         Height          =   400
         Left            =   2100
         TabIndex        =   79
         Text            =   "0"
         Top             =   2900
         Width           =   600
      End
      Begin VB.TextBox Text36 
         Height          =   400
         Left            =   1100
         TabIndex        =   78
         Text            =   "0"
         Top             =   2900
         Width           =   600
      End
      Begin VB.TextBox Text35 
         Height          =   400
         Left            =   4100
         TabIndex        =   77
         Text            =   "0"
         Top             =   2300
         Width           =   600
      End
      Begin VB.TextBox Text34 
         Height          =   400
         Left            =   3100
         TabIndex        =   76
         Text            =   "0"
         Top             =   2300
         Width           =   600
      End
      Begin VB.TextBox Text33 
         Height          =   400
         Left            =   2100
         TabIndex        =   75
         Text            =   "0"
         Top             =   2300
         Width           =   600
      End
      Begin VB.TextBox Text32 
         Height          =   400
         Left            =   1100
         TabIndex        =   74
         Text            =   "0"
         Top             =   2300
         Width           =   600
      End
      Begin VB.TextBox Text31 
         Height          =   400
         Left            =   4100
         TabIndex        =   66
         Text            =   "0"
         Top             =   1300
         Width           =   600
      End
      Begin VB.TextBox Text30 
         Height          =   400
         Left            =   3100
         TabIndex        =   65
         Text            =   "0"
         Top             =   1300
         Width           =   600
      End
      Begin VB.TextBox Text29 
         Height          =   400
         Left            =   2100
         TabIndex        =   64
         Text            =   "0"
         Top             =   1300
         Width           =   600
      End
      Begin VB.TextBox Text28 
         Height          =   400
         Left            =   1100
         TabIndex        =   63
         Text            =   "0"
         Top             =   1300
         Width           =   600
      End
      Begin VB.Label Label41 
         BackColor       =   &H0000C0C0&
         Caption         =   "Четв"
         Height          =   300
         Left            =   400
         TabIndex        =   73
         Top             =   5300
         Width           =   600
      End
      Begin VB.Label Label40 
         BackColor       =   &H0000C0C0&
         Caption         =   "Трет"
         Height          =   300
         Left            =   400
         TabIndex        =   72
         Top             =   4700
         Width           =   600
      End
      Begin VB.Label Label39 
         BackColor       =   &H0000C0C0&
         Caption         =   "Втор"
         Height          =   300
         Left            =   400
         TabIndex        =   71
         Top             =   4100
         Width           =   600
      End
      Begin VB.Label Label38 
         BackColor       =   &H0000C0C0&
         Caption         =   "Перв"
         Height          =   300
         Left            =   400
         TabIndex        =   70
         Top             =   3500
         Width           =   600
      End
      Begin VB.Label Label37 
         BackColor       =   &H0000C0C0&
         Caption         =   "Умен"
         Height          =   300
         Left            =   400
         TabIndex        =   69
         Top             =   2900
         Width           =   600
      End
      Begin VB.Label Label36 
         BackColor       =   &H0000C0C0&
         Caption         =   "Полн"
         Height          =   300
         Left            =   400
         TabIndex        =   68
         Top             =   2300
         Width           =   600
      End
      Begin VB.Label Label35 
         BackColor       =   &H0000C0C0&
         Caption         =   "       Потеря нач. скорости Vо"
         Height          =   300
         Left            =   1100
         TabIndex        =   67
         Top             =   1900
         Width           =   3600
      End
      Begin VB.Label Label34 
         BackColor       =   &H0000C0C0&
         Caption         =   "4 Бат"
         Height          =   300
         Left            =   4100
         TabIndex        =   62
         Top             =   500
         Width           =   600
      End
      Begin VB.Label Label33 
         BackColor       =   &H0000C0C0&
         Caption         =   "3 Бат"
         Height          =   300
         Left            =   3100
         TabIndex        =   61
         Top             =   500
         Width           =   600
      End
      Begin VB.Label Label32 
         BackColor       =   &H0000C0C0&
         Caption         =   "2 Бат"
         Height          =   300
         Left            =   2100
         TabIndex        =   60
         Top             =   500
         Width           =   600
      End
      Begin VB.Label Label31 
         BackColor       =   &H0000C0C0&
         Caption         =   "1 Бат"
         Height          =   300
         Left            =   1100
         TabIndex        =   59
         Top             =   500
         Width           =   600
      End
      Begin VB.Label Label30 
         BackColor       =   &H0000C0C0&
         Caption         =   "            Температура заряда"
         Height          =   300
         Left            =   1100
         TabIndex        =   58
         Top             =   900
         Width           =   3600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      Caption         =   " Боевой порядок"
      Height          =   8055
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   6700
      Begin VB.TextBox Text27 
         Height          =   400
         Left            =   5200
         TabIndex        =   56
         Text            =   "0"
         Top             =   7100
         Width           =   1000
      End
      Begin VB.TextBox Text26 
         Height          =   400
         Left            =   3200
         TabIndex        =   54
         Text            =   "0"
         Top             =   7100
         Width           =   1500
      End
      Begin VB.TextBox Text25 
         Height          =   400
         Left            =   1200
         TabIndex        =   52
         Text            =   "0"
         Top             =   7100
         Width           =   1500
      End
      Begin VB.TextBox Text24 
         Height          =   400
         Left            =   5200
         TabIndex        =   50
         Text            =   "0"
         Top             =   6400
         Width           =   1000
      End
      Begin VB.TextBox Text23 
         Height          =   400
         Left            =   3200
         TabIndex        =   48
         Text            =   "0"
         Top             =   6400
         Width           =   1500
      End
      Begin VB.TextBox Text22 
         Height          =   400
         Left            =   1200
         TabIndex        =   46
         Text            =   "0"
         Top             =   6400
         Width           =   1500
      End
      Begin VB.TextBox Text21 
         Height          =   400
         Left            =   5200
         TabIndex        =   44
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox Text20 
         Height          =   400
         Left            =   3200
         TabIndex        =   42
         Text            =   "0"
         Top             =   5700
         Width           =   1500
      End
      Begin VB.TextBox Text19 
         Height          =   400
         Left            =   1200
         TabIndex        =   40
         Text            =   "0"
         Top             =   5700
         Width           =   1500
      End
      Begin VB.TextBox Text18 
         Height          =   400
         Left            =   5200
         TabIndex        =   38
         Text            =   "0"
         Top             =   5000
         Width           =   1000
      End
      Begin VB.TextBox Text17 
         Height          =   400
         Left            =   3200
         TabIndex        =   36
         Text            =   "0"
         Top             =   5000
         Width           =   1500
      End
      Begin VB.TextBox Text16 
         Height          =   400
         Left            =   1200
         TabIndex        =   34
         Text            =   "0"
         Top             =   5000
         Width           =   1500
      End
      Begin VB.TextBox Text15 
         Height          =   400
         Left            =   5200
         TabIndex        =   32
         Text            =   "0"
         Top             =   4300
         Width           =   1000
      End
      Begin VB.TextBox Text14 
         Height          =   400
         Left            =   3200
         TabIndex        =   30
         Text            =   "0"
         Top             =   4300
         Width           =   1500
      End
      Begin VB.TextBox Text13 
         Height          =   400
         Left            =   1200
         TabIndex        =   29
         Text            =   "0"
         Top             =   4300
         Width           =   1500
      End
      Begin VB.TextBox Text12 
         Height          =   400
         Left            =   5200
         TabIndex        =   28
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox Text11 
         Height          =   400
         Left            =   3200
         TabIndex        =   26
         Text            =   "0"
         Top             =   3100
         Width           =   1500
      End
      Begin VB.TextBox Text10 
         Height          =   400
         Left            =   1200
         TabIndex        =   24
         Text            =   "0"
         Top             =   3100
         Width           =   1500
      End
      Begin VB.TextBox Text9 
         Height          =   400
         Left            =   5200
         TabIndex        =   22
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox Text8 
         Height          =   400
         Left            =   3200
         TabIndex        =   20
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox Text7 
         Height          =   400
         Left            =   1200
         TabIndex        =   18
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox Text6 
         Height          =   400
         Left            =   5200
         TabIndex        =   13
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox Text5 
         Height          =   400
         Left            =   3200
         TabIndex        =   11
         Text            =   "0"
         Top             =   1700
         Width           =   1500
      End
      Begin VB.TextBox Text4 
         Height          =   400
         Left            =   1200
         TabIndex        =   9
         Text            =   "0"
         Top             =   1700
         Width           =   1500
      End
      Begin VB.TextBox Text3 
         Height          =   400
         Left            =   5200
         TabIndex        =   7
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox Text2 
         Height          =   400
         Left            =   3200
         TabIndex        =   5
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         Height          =   400
         Left            =   1200
         TabIndex        =   3
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.Label Label29 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   55
         Top             =   7100
         Width           =   300
      End
      Begin VB.Label Label28 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2800
         TabIndex        =   53
         Top             =   7100
         Width           =   300
      End
      Begin VB.Label Label27 
         BackColor       =   &H0000C0C0&
         Caption         =   "КНП 5 Х="
         Height          =   300
         Left            =   100
         TabIndex        =   51
         Top             =   7100
         Width           =   1100
      End
      Begin VB.Label Label26 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   49
         Top             =   6400
         Width           =   300
      End
      Begin VB.Label Label25 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2800
         TabIndex        =   47
         Top             =   6400
         Width           =   300
      End
      Begin VB.Label Label24 
         BackColor       =   &H0000C0C0&
         Caption         =   "КНП 4 Х="
         Height          =   300
         Left            =   120
         TabIndex        =   45
         Top             =   6405
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   43
         Top             =   5700
         Width           =   300
      End
      Begin VB.Label Label21 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2800
         TabIndex        =   41
         Top             =   5700
         Width           =   300
      End
      Begin VB.Label Label15 
         BackColor       =   &H0000C0C0&
         Caption         =   "КНП 3 Х="
         Height          =   300
         Left            =   100
         TabIndex        =   39
         Top             =   5700
         Width           =   1100
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   37
         Top             =   5000
         Width           =   300
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2800
         TabIndex        =   35
         Top             =   5000
         Width           =   300
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000C0C0&
         Caption         =   "КНП 2 Х="
         Height          =   300
         Left            =   100
         TabIndex        =   33
         Top             =   5000
         Width           =   1100
      End
      Begin VB.Label Label22 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   31
         Top             =   4300
         Width           =   300
      End
      Begin VB.Label Label20 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   27
         Top             =   3100
         Width           =   300
      End
      Begin VB.Label Label19 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2800
         TabIndex        =   25
         Top             =   3100
         Width           =   300
      End
      Begin VB.Label Label18 
         BackColor       =   &H0000C0C0&
         Caption         =   "4 Бат Х="
         Height          =   300
         Left            =   200
         TabIndex        =   23
         Top             =   3100
         Width           =   1000
      End
      Begin VB.Label Label17 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   21
         Top             =   2400
         Width           =   300
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2800
         TabIndex        =   19
         Top             =   2400
         Width           =   300
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2800
         TabIndex        =   17
         Top             =   4300
         Width           =   300
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000C0C0&
         Caption         =   "КНП 1 Х="
         Height          =   300
         Left            =   100
         TabIndex        =   16
         Top             =   4300
         Width           =   1100
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000C0C0&
         Caption         =   "                                               КНП"
         Height          =   300
         Left            =   300
         TabIndex        =   15
         Top             =   3800
         Width           =   6200
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000C0C0&
         Caption         =   "3 Бат Х="
         Height          =   300
         Left            =   200
         TabIndex        =   14
         Top             =   2400
         Width           =   1000
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   12
         Top             =   1700
         Width           =   300
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   400
         Left            =   2800
         TabIndex        =   10
         Top             =   1700
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C0C0&
         Caption         =   "2 Бат Х="
         Height          =   300
         Left            =   200
         TabIndex        =   8
         Top             =   1700
         Width           =   1000
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   4800
         TabIndex        =   6
         Top             =   1005
         Width           =   300
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         Caption         =   "У="
         Height          =   405
         Left            =   2800
         TabIndex        =   4
         Top             =   1005
         Width           =   300
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "1 Бат Х="
         Height          =   300
         Left            =   200
         TabIndex        =   2
         Top             =   1000
         Width           =   1000
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "                                         ОГНЕВЫЕ"
         Height          =   300
         Left            =   300
         TabIndex        =   1
         Top             =   480
         Width           =   6200
      End
   End
End
Attribute VB_Name = "BP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
