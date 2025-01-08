VERSION 5.00
Begin VB.Form BP 
   Caption         =   "Боевой порядок"
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
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "ВЫХОД"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   17600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   8300
      Width           =   2500
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "Добавить ОП, НП"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3400
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   8300
      Width           =   3000
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "КАРТА"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   100
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   8300
      Width           =   3000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Добавить в Архив"
      Height          =   700
      Left            =   17600
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   7440
      Width           =   2500
   End
   Begin VB.CommandButton PosmNP 
      BackColor       =   &H0080FF80&
      Caption         =   "Посмотреть Архив "
      Height          =   700
      Left            =   14000
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   7440
      Width           =   2500
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Выбор плановых ОП"
      Height          =   2200
      Left            =   9000
      TabIndex        =   156
      Top             =   5950
      Width           =   4700
      Begin VB.ComboBox pplOP3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   3100
         TabIndex        =   167
         Text            =   "0"
         Top             =   1000
         Width           =   1400
      End
      Begin VB.ComboBox pplOP2 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   450
         Left            =   1600
         TabIndex        =   166
         Text            =   "0"
         Top             =   1000
         Width           =   1400
      End
      Begin VB.ComboBox pplOP1 
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
         Left            =   100
         TabIndex        =   165
         Text            =   "0"
         Top             =   1000
         Width           =   1400
      End
      Begin VB.Label Label65 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   3 Бат"
         Height          =   300
         Left            =   3200
         TabIndex        =   159
         Top             =   400
         Width           =   1000
      End
      Begin VB.Label Label64 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   2 Бат"
         Height          =   300
         Left            =   1700
         TabIndex        =   158
         Top             =   400
         Width           =   1000
      End
      Begin VB.Label Label63 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   1 Бат"
         Height          =   300
         Left            =   200
         TabIndex        =   157
         Top             =   400
         Width           =   1000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Метеоданные"
      Height          =   7000
      Left            =   14000
      TabIndex        =   90
      Top             =   100
      Width           =   6135
      Begin VB.CommandButton polMeteo 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Получить Метеосредний"
         Height          =   855
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   5500
         Width           =   2000
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Составить Метеоприближенный"
         Height          =   855
         Left            =   100
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   5500
         Width           =   2900
      End
      Begin VB.TextBox pW60 
         Height          =   400
         Left            =   4900
         TabIndex        =   152
         Text            =   "0"
         Top             =   4800
         Width           =   700
      End
      Begin VB.TextBox pW50 
         Height          =   400
         Left            =   4900
         TabIndex        =   151
         Text            =   "0"
         Top             =   4400
         Width           =   700
      End
      Begin VB.TextBox pAw60 
         Height          =   400
         Left            =   4100
         TabIndex        =   150
         Text            =   "0"
         Top             =   4800
         Width           =   700
      End
      Begin VB.TextBox pAw50 
         Height          =   400
         Left            =   4100
         TabIndex        =   149
         Text            =   "0"
         Top             =   4400
         Width           =   700
      End
      Begin VB.TextBox pDt60 
         Height          =   400
         Left            =   3300
         TabIndex        =   148
         Text            =   "0"
         Top             =   4800
         Width           =   700
      End
      Begin VB.TextBox pdT50 
         Height          =   400
         Left            =   3300
         TabIndex        =   147
         Text            =   "0"
         Top             =   4400
         Width           =   700
      End
      Begin VB.TextBox pW40 
         Height          =   400
         Left            =   4900
         TabIndex        =   144
         Text            =   "0"
         Top             =   4000
         Width           =   700
      End
      Begin VB.TextBox pW30 
         Height          =   400
         Left            =   4900
         TabIndex        =   143
         Text            =   "0"
         Top             =   3600
         Width           =   700
      End
      Begin VB.TextBox pW24 
         Height          =   400
         Left            =   4900
         TabIndex        =   142
         Text            =   "0"
         Top             =   3200
         Width           =   700
      End
      Begin VB.TextBox pW20 
         Height          =   400
         Left            =   4900
         TabIndex        =   141
         Text            =   "0"
         Top             =   2800
         Width           =   700
      End
      Begin VB.TextBox pW16 
         Height          =   400
         Left            =   4900
         TabIndex        =   140
         Text            =   "0"
         Top             =   2400
         Width           =   700
      End
      Begin VB.TextBox pW12 
         Height          =   400
         Left            =   4900
         TabIndex        =   139
         Text            =   "0"
         Top             =   2000
         Width           =   700
      End
      Begin VB.TextBox pW08 
         Height          =   400
         Left            =   4900
         TabIndex        =   138
         Text            =   "0"
         Top             =   1600
         Width           =   700
      End
      Begin VB.TextBox pW04 
         Height          =   400
         Left            =   4900
         TabIndex        =   137
         Text            =   "0"
         Top             =   1200
         Width           =   700
      End
      Begin VB.TextBox pW02 
         Height          =   400
         Left            =   4900
         TabIndex        =   136
         Text            =   "0"
         Top             =   800
         Width           =   700
      End
      Begin VB.TextBox pAw40 
         Height          =   400
         Left            =   4100
         TabIndex        =   135
         Text            =   "0"
         Top             =   4000
         Width           =   700
      End
      Begin VB.TextBox pAw30 
         Height          =   400
         Left            =   4100
         TabIndex        =   134
         Text            =   "0"
         Top             =   3600
         Width           =   700
      End
      Begin VB.TextBox pAw24 
         Height          =   400
         Left            =   4100
         TabIndex        =   133
         Text            =   "0"
         Top             =   3200
         Width           =   700
      End
      Begin VB.TextBox pAw20 
         Height          =   400
         Left            =   4100
         TabIndex        =   132
         Text            =   "0"
         Top             =   2800
         Width           =   700
      End
      Begin VB.TextBox pAw16 
         Height          =   400
         Left            =   4100
         TabIndex        =   131
         Text            =   "0"
         Top             =   2400
         Width           =   700
      End
      Begin VB.TextBox pAw12 
         Height          =   400
         Left            =   4100
         TabIndex        =   130
         Text            =   "0"
         Top             =   2000
         Width           =   700
      End
      Begin VB.TextBox pAw08 
         Height          =   400
         Left            =   4100
         TabIndex        =   129
         Text            =   "0"
         Top             =   1600
         Width           =   700
      End
      Begin VB.TextBox pAw04 
         Height          =   400
         Left            =   4100
         TabIndex        =   128
         Text            =   "0"
         Top             =   1200
         Width           =   700
      End
      Begin VB.TextBox pAw02 
         Height          =   400
         Left            =   4100
         TabIndex        =   127
         Text            =   "0"
         Top             =   800
         Width           =   700
      End
      Begin VB.TextBox pdT40 
         Height          =   400
         Left            =   3300
         TabIndex        =   126
         Text            =   "0"
         Top             =   4000
         Width           =   700
      End
      Begin VB.TextBox pdT30 
         Height          =   400
         Left            =   3300
         TabIndex        =   125
         Text            =   "0"
         Top             =   3600
         Width           =   700
      End
      Begin VB.TextBox pdT24 
         Height          =   400
         Left            =   3300
         TabIndex        =   124
         Text            =   "0"
         Top             =   3200
         Width           =   700
      End
      Begin VB.TextBox pdT20 
         Height          =   400
         Left            =   3300
         TabIndex        =   123
         Text            =   "0"
         Top             =   2800
         Width           =   700
      End
      Begin VB.TextBox pdT16 
         Height          =   400
         Left            =   3300
         TabIndex        =   122
         Text            =   "0"
         Top             =   2400
         Width           =   700
      End
      Begin VB.TextBox pdT12 
         Height          =   400
         Left            =   3300
         TabIndex        =   121
         Text            =   "0"
         Top             =   2000
         Width           =   700
      End
      Begin VB.TextBox pdT08 
         Height          =   400
         Left            =   3300
         TabIndex        =   120
         Text            =   "0"
         Top             =   1600
         Width           =   700
      End
      Begin VB.TextBox pdT04 
         Height          =   400
         Left            =   3300
         TabIndex        =   119
         Text            =   "0"
         Top             =   1200
         Width           =   700
      End
      Begin VB.TextBox pdT02 
         Height          =   400
         Left            =   3300
         TabIndex        =   118
         Text            =   "0"
         Top             =   800
         Width           =   700
      End
      Begin VB.TextBox pW 
         Height          =   400
         Left            =   1200
         TabIndex        =   105
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pAw 
         Height          =   400
         Left            =   1200
         TabIndex        =   104
         Text            =   "0"
         Top             =   2500
         Width           =   1000
      End
      Begin VB.TextBox pTv 
         Height          =   400
         Left            =   1200
         TabIndex        =   103
         Text            =   "15"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox pH 
         Height          =   400
         Left            =   1200
         TabIndex        =   102
         Text            =   "750"
         Top             =   1100
         Width           =   1000
      End
      Begin VB.TextBox phmet 
         Height          =   400
         Left            =   1200
         TabIndex        =   101
         Text            =   "0"
         Top             =   400
         Width           =   1000
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Height          =   1400
         Left            =   100
         TabIndex        =   96
         Top             =   3900
         Width           =   2100
         Begin VB.OptionButton pVR2 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1200
            TabIndex        =   100
            Top             =   800
            Width           =   615
         End
         Begin VB.OptionButton pDMK 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   360
            TabIndex        =   99
            Top             =   800
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.Label Label47 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ВР-2"
            Height          =   300
            Left            =   1200
            TabIndex        =   98
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label46 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ДМК"
            Height          =   300
            Left            =   300
            TabIndex        =   97
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.Label Label68 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Угол ветра в ДУ"
         Height          =   375
         Left            =   100
         TabIndex        =   153
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label61 
         BackColor       =   &H00C0C0C0&
         Caption         =   "60"
         Height          =   300
         Left            =   2700
         TabIndex        =   146
         Top             =   4800
         Width           =   500
      End
      Begin VB.Label Label60 
         BackColor       =   &H00C0C0C0&
         Caption         =   "50"
         Height          =   300
         Left            =   2700
         TabIndex        =   145
         Top             =   4400
         Width           =   500
      End
      Begin VB.Label Label59 
         BackColor       =   &H00C0C0C0&
         Caption         =   "40"
         Height          =   300
         Left            =   2700
         TabIndex        =   117
         Top             =   4000
         Width           =   500
      End
      Begin VB.Label Label58 
         BackColor       =   &H00C0C0C0&
         Caption         =   "30"
         Height          =   300
         Left            =   2700
         TabIndex        =   116
         Top             =   3600
         Width           =   500
      End
      Begin VB.Label Label57 
         BackColor       =   &H00C0C0C0&
         Caption         =   "24"
         Height          =   300
         Left            =   2700
         TabIndex        =   115
         Top             =   3200
         Width           =   500
      End
      Begin VB.Label Label56 
         BackColor       =   &H00C0C0C0&
         Caption         =   "20"
         Height          =   300
         Left            =   2700
         TabIndex        =   114
         Top             =   2800
         Width           =   500
      End
      Begin VB.Label Label55 
         BackColor       =   &H00C0C0C0&
         Caption         =   "16"
         Height          =   300
         Left            =   2700
         TabIndex        =   113
         Top             =   2400
         Width           =   500
      End
      Begin VB.Label Label54 
         BackColor       =   &H00C0C0C0&
         Caption         =   "12"
         Height          =   300
         Left            =   2700
         TabIndex        =   112
         Top             =   2000
         Width           =   500
      End
      Begin VB.Label Label53 
         BackColor       =   &H00C0C0C0&
         Caption         =   "08"
         Height          =   300
         Left            =   2700
         TabIndex        =   111
         Top             =   1600
         Width           =   500
      End
      Begin VB.Label Label52 
         BackColor       =   &H00C0C0C0&
         Caption         =   "04"
         Height          =   300
         Left            =   2700
         TabIndex        =   110
         Top             =   1200
         Width           =   500
      End
      Begin VB.Label Label51 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  W"
         Height          =   300
         Left            =   5000
         TabIndex        =   109
         Top             =   400
         Width           =   500
      End
      Begin VB.Label Label50 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Aw"
         Height          =   300
         Left            =   4200
         TabIndex        =   108
         Top             =   400
         Width           =   500
      End
      Begin VB.Label Label49 
         BackColor       =   &H00C0C0C0&
         Caption         =   " dT"
         Height          =   300
         Left            =   3400
         TabIndex        =   107
         Top             =   400
         Width           =   500
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         Caption         =   "02"
         Height          =   300
         Left            =   2700
         TabIndex        =   106
         Top             =   800
         Width           =   500
      End
      Begin VB.Label Label45 
         BackColor       =   &H00C0C0C0&
         Caption         =   "W="
         Height          =   300
         Left            =   100
         TabIndex        =   95
         Top             =   3200
         Width           =   1000
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aw="
         Height          =   300
         Left            =   100
         TabIndex        =   94
         Top             =   2500
         Width           =   500
      End
      Begin VB.Label Label43 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tv="
         Height          =   300
         Left            =   100
         TabIndex        =   93
         Top             =   1800
         Width           =   500
      End
      Begin VB.Label Label42 
         BackColor       =   &H00C0C0C0&
         Caption         =   "H="
         Height          =   300
         Left            =   100
         TabIndex        =   92
         Top             =   1100
         Width           =   500
      End
      Begin VB.Label Label41 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h Метео"
         Height          =   300
         Left            =   100
         TabIndex        =   91
         Top             =   400
         Width           =   1000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Балистика"
      Height          =   5700
      Left            =   9000
      TabIndex        =   57
      Top             =   100
      Width           =   4695
      Begin VB.TextBox pV034 
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   3300
         TabIndex        =   89
         Text            =   "-0,8"
         Top             =   5000
         Width           =   700
      End
      Begin VB.TextBox pV033 
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   3300
         TabIndex        =   88
         Text            =   "-0,5"
         Top             =   4500
         Width           =   700
      End
      Begin VB.TextBox pV032 
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   3300
         TabIndex        =   87
         Text            =   "-0,5"
         Top             =   4000
         Width           =   700
      End
      Begin VB.TextBox pV031 
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   3300
         TabIndex        =   86
         Text            =   "-0,5"
         Top             =   3500
         Width           =   700
      End
      Begin VB.TextBox pV03Y 
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   3300
         TabIndex        =   85
         Text            =   "-0,5"
         Top             =   3000
         Width           =   700
      End
      Begin VB.TextBox pV03p 
         ForeColor       =   &H00FF0000&
         Height          =   400
         Left            =   3300
         TabIndex        =   84
         Text            =   "0,5"
         Top             =   2500
         Width           =   700
      End
      Begin VB.TextBox pV024 
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   2200
         TabIndex        =   83
         Text            =   "-0,8"
         Top             =   5000
         Width           =   700
      End
      Begin VB.TextBox pV023 
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   2200
         TabIndex        =   82
         Text            =   "-0,5"
         Top             =   4500
         Width           =   700
      End
      Begin VB.TextBox pV022 
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   2200
         TabIndex        =   81
         Text            =   "-0,5"
         Top             =   4000
         Width           =   700
      End
      Begin VB.TextBox pV021 
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   2200
         TabIndex        =   80
         Text            =   "-0,5"
         Top             =   3500
         Width           =   700
      End
      Begin VB.TextBox pV02y 
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   2200
         TabIndex        =   79
         Text            =   "-0,5"
         Top             =   3000
         Width           =   700
      End
      Begin VB.TextBox pV02p 
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   2200
         TabIndex        =   78
         Text            =   "0,5"
         Top             =   2500
         Width           =   700
      End
      Begin VB.TextBox pV014 
         Height          =   400
         Left            =   1100
         TabIndex        =   77
         Text            =   "-0,8"
         Top             =   5000
         Width           =   700
      End
      Begin VB.TextBox pV013 
         Height          =   400
         Left            =   1100
         TabIndex        =   76
         Text            =   "-0,5"
         Top             =   4500
         Width           =   700
      End
      Begin VB.TextBox pV012 
         Height          =   400
         Left            =   1100
         TabIndex        =   75
         Text            =   "-0,5"
         Top             =   4000
         Width           =   700
      End
      Begin VB.TextBox pV011 
         Height          =   400
         Left            =   1100
         TabIndex        =   74
         Text            =   "-0,5"
         Top             =   3500
         Width           =   700
      End
      Begin VB.TextBox pV01y 
         Height          =   400
         Left            =   1100
         TabIndex        =   73
         Text            =   "-0,5"
         Top             =   3000
         Width           =   700
      End
      Begin VB.TextBox pV01p 
         Height          =   400
         Left            =   1100
         TabIndex        =   72
         Text            =   "0,5"
         Top             =   2500
         Width           =   700
      End
      Begin VB.TextBox pTz3 
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   3300
         TabIndex        =   64
         Text            =   "15"
         Top             =   1100
         Width           =   700
      End
      Begin VB.TextBox pTz2 
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   2200
         TabIndex        =   63
         Text            =   "15"
         Top             =   1100
         Width           =   700
      End
      Begin VB.TextBox pTz1 
         Height          =   405
         Left            =   1100
         TabIndex        =   62
         Text            =   "15"
         Top             =   1100
         Width           =   700
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Четверт"
         Height          =   300
         Left            =   100
         TabIndex        =   71
         Top             =   5000
         Width           =   1000
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Трет"
         Height          =   300
         Left            =   100
         TabIndex        =   70
         Top             =   4500
         Width           =   1000
      End
      Begin VB.Label Label38 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Втор"
         Height          =   300
         Left            =   100
         TabIndex        =   69
         Top             =   4000
         Width           =   1000
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Прев"
         Height          =   300
         Left            =   100
         TabIndex        =   68
         Top             =   3500
         Width           =   1000
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Умен"
         Height          =   300
         Left            =   100
         TabIndex        =   67
         Top             =   3000
         Width           =   1000
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Полн"
         Height          =   300
         Left            =   100
         TabIndex        =   66
         Top             =   2500
         Width           =   1000
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C0C0&
         Caption         =   "          Потеря нач скорости"
         Height          =   375
         Left            =   300
         TabIndex        =   65
         Top             =   1800
         Width           =   3700
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Тз="
         Height          =   300
         Left            =   100
         TabIndex        =   61
         Top             =   1100
         Width           =   400
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 Бат"
         Height          =   300
         Left            =   3300
         TabIndex        =   60
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 Бат"
         Height          =   300
         Left            =   2200
         TabIndex        =   59
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 Бат"
         Height          =   300
         Left            =   1100
         TabIndex        =   58
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Боевой порядок"
      Height          =   8055
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   8655
      Begin VB.ComboBox pplNP5 
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
         Left            =   6700
         TabIndex        =   172
         Text            =   "0"
         Top             =   7000
         Width           =   1800
      End
      Begin VB.ComboBox pplNP4 
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
         Left            =   6700
         TabIndex        =   171
         Text            =   "0"
         Top             =   6300
         Width           =   1800
      End
      Begin VB.ComboBox pplNP3 
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
         Left            =   6700
         TabIndex        =   170
         Text            =   "0"
         Top             =   5600
         Width           =   1800
      End
      Begin VB.ComboBox pplNP2 
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
         Left            =   6700
         TabIndex        =   169
         Text            =   "0"
         Top             =   4900
         Width           =   1800
      End
      Begin VB.ComboBox pplNP1 
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
         Left            =   6700
         TabIndex        =   168
         Text            =   "0"
         Top             =   4200
         Width           =   1800
      End
      Begin VB.TextBox phkp5 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5500
         TabIndex        =   55
         Text            =   "0"
         Top             =   7000
         Width           =   1000
      End
      Begin VB.TextBox phkp4 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5500
         TabIndex        =   54
         Text            =   "0"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.TextBox phkp3 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5500
         TabIndex        =   53
         Text            =   "0"
         Top             =   5600
         Width           =   1000
      End
      Begin VB.TextBox phkp2 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5500
         TabIndex        =   52
         Text            =   "0"
         Top             =   4900
         Width           =   1000
      End
      Begin VB.TextBox phkp1 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5500
         TabIndex        =   51
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pYkp5 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3300
         TabIndex        =   45
         Text            =   "0"
         Top             =   7000
         Width           =   1500
      End
      Begin VB.TextBox pYkp4 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3300
         TabIndex        =   44
         Text            =   "0"
         Top             =   6300
         Width           =   1500
      End
      Begin VB.TextBox pYkp3 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3300
         TabIndex        =   43
         Text            =   "0"
         Top             =   5600
         Width           =   1500
      End
      Begin VB.TextBox pYkp2 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3300
         TabIndex        =   42
         Text            =   "0"
         Top             =   4900
         Width           =   1500
      End
      Begin VB.TextBox pYkp1 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3300
         TabIndex        =   41
         Text            =   "0"
         Top             =   4200
         Width           =   1500
      End
      Begin VB.TextBox pXkp5 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1100
         TabIndex        =   36
         Text            =   "0"
         Top             =   7000
         Width           =   1500
      End
      Begin VB.TextBox pXkp4 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1100
         TabIndex        =   35
         Text            =   "0"
         Top             =   6300
         Width           =   1500
      End
      Begin VB.TextBox pXkp3 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1100
         TabIndex        =   34
         Text            =   "0"
         Top             =   5600
         Width           =   1500
      End
      Begin VB.TextBox pXkp2 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1100
         TabIndex        =   33
         Text            =   "0"
         Top             =   4900
         Width           =   1500
      End
      Begin VB.TextBox pXkp1 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1100
         TabIndex        =   32
         Text            =   "0"
         Top             =   4200
         Width           =   1500
      End
      Begin VB.TextBox pOH3 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   7300
         TabIndex        =   25
         Text            =   "0"
         Top             =   2500
         Width           =   1000
      End
      Begin VB.TextBox pOH2 
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
         Height          =   495
         Left            =   7300
         TabIndex        =   24
         Text            =   "0"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox pOH1 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7300
         TabIndex        =   23
         Text            =   "0"
         Top             =   1100
         Width           =   1000
      End
      Begin VB.TextBox ph3 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5500
         TabIndex        =   19
         Text            =   "0"
         Top             =   2500
         Width           =   1000
      End
      Begin VB.TextBox ph2 
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
         Height          =   495
         Left            =   5500
         TabIndex        =   18
         Text            =   "0"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox ph1 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5500
         TabIndex        =   17
         Text            =   "0"
         Top             =   1100
         Width           =   1000
      End
      Begin VB.TextBox pY3 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3300
         TabIndex        =   13
         Text            =   "0"
         Top             =   2500
         Width           =   1500
      End
      Begin VB.TextBox pY2 
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
         Height          =   495
         Left            =   3300
         TabIndex        =   12
         Text            =   "0"
         Top             =   1800
         Width           =   1500
      End
      Begin VB.TextBox pY1 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3300
         TabIndex        =   11
         Text            =   "0"
         Top             =   1100
         Width           =   1500
      End
      Begin VB.TextBox pX3 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1100
         TabIndex        =   7
         Text            =   "0"
         Top             =   2500
         Width           =   1500
      End
      Begin VB.TextBox pX2 
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
         Height          =   495
         Left            =   1100
         TabIndex        =   6
         Text            =   "0"
         Top             =   1800
         Width           =   1500
      End
      Begin VB.TextBox pX1 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1100
         TabIndex        =   5
         Text            =   "0"
         Top             =   1100
         Width           =   1500
      End
      Begin VB.Label Label62 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Плановые НП"
         Height          =   375
         Left            =   6900
         TabIndex        =   155
         Top             =   3500
         Width           =   1600
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   2800
         TabIndex        =   56
         Top             =   7000
         Width           =   400
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   5000
         TabIndex        =   50
         Top             =   7000
         Width           =   400
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   5000
         TabIndex        =   49
         Top             =   6300
         Width           =   400
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   5000
         TabIndex        =   48
         Top             =   5600
         Width           =   400
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   5000
         TabIndex        =   47
         Top             =   4900
         Width           =   400
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   5000
         TabIndex        =   46
         Top             =   4200
         Width           =   400
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   2800
         TabIndex        =   40
         Top             =   6300
         Width           =   400
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   2800
         TabIndex        =   39
         Top             =   5600
         Width           =   400
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   2800
         TabIndex        =   38
         Top             =   4900
         Width           =   400
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   2800
         TabIndex        =   37
         Top             =   4200
         Width           =   400
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5 КНП="
         Height          =   300
         Left            =   100
         TabIndex        =   31
         Top             =   7000
         Width           =   1000
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4КНП="
         Height          =   300
         Left            =   100
         TabIndex        =   30
         Top             =   6300
         Width           =   1000
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3КНП="
         Height          =   300
         Left            =   100
         TabIndex        =   29
         Top             =   5600
         Width           =   1000
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 КНП="
         Height          =   300
         Left            =   100
         TabIndex        =   28
         Top             =   4900
         Width           =   1000
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 КНП="
         Height          =   300
         Left            =   100
         TabIndex        =   27
         Top             =   4200
         Width           =   1000
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "                                                 КНП"
         Height          =   300
         Left            =   100
         TabIndex        =   26
         Top             =   3500
         Width           =   6400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OH="
         Height          =   400
         Left            =   6700
         TabIndex        =   22
         Top             =   2500
         Width           =   500
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OH="
         Height          =   400
         Left            =   6700
         TabIndex        =   21
         Top             =   1800
         Width           =   500
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OH="
         Height          =   400
         Left            =   6700
         TabIndex        =   20
         Top             =   1100
         Width           =   500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   5000
         TabIndex        =   16
         Top             =   2500
         Width           =   300
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   5000
         TabIndex        =   15
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         Height          =   300
         Left            =   5000
         TabIndex        =   14
         Top             =   1100
         Width           =   300
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   2800
         TabIndex        =   10
         Top             =   2500
         Width           =   400
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   2800
         TabIndex        =   9
         Top             =   1800
         Width           =   400
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         Height          =   300
         Left            =   2800
         TabIndex        =   8
         Top             =   1100
         Width           =   400
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 Бат Х="
         Height          =   300
         Left            =   100
         TabIndex        =   4
         Top             =   2500
         Width           =   1000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 Бат Х="
         Height          =   300
         Left            =   100
         TabIndex        =   3
         Top             =   1800
         Width           =   1000
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 Бат Х="
         Height          =   300
         Left            =   100
         TabIndex        =   2
         Top             =   1100
         Width           =   1000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "                                                           ОГНЕВЫЕ"
         Height          =   300
         Left            =   105
         TabIndex        =   1
         Top             =   405
         Width           =   8200
      End
   End
End
Attribute VB_Name = "BP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vvod, W, Dsnos, w200, w400, w800, w1200, w1600, w2000, w2400, w3000, w4000 As Single
Public hmet, Tv, h, Aw As Single

Private Sub BP_s_Faila_Click()
Dim t(5, 9) As String
Dim stroka As String
Dim myArr() As String
Dim i As Integer, j As Integer
'записываем файл в масив
Open "D:\YO_NA\BP_Brig\bp" For Input As #1
10: If EOF(1) Then GoTo 20
  Line Input #1, stroka
   myArr = Split(stroka, ",")
  For i = 0 To 9
        t(j, i) = myArr(i)
        Next i
        j = j + 1
   GoTo 10
20: Close #1
'пердаем данные с масива в поля батарей
pX1.Text = Val(t(0, 3)): pY1.Text = Val(t(1, 3)): ph1.Text = Val(t(2, 3)): pOH1.Text = Val(t(3, 3)): pTz1.Text = Val(t(4, 3))
pX2.Text = Val(t(0, 4)): pY2.Text = Val(t(1, 4)): ph2.Text = Val(t(2, 4)): pOH2.Text = Val(t(3, 4)): pTz2.Text = Val(t(4, 4))
pX3.Text = Val(t(0, 5)): pY3.Text = Val(t(1, 5)): ph3.Text = Val(t(2, 5)): pOH3.Text = Val(t(3, 5)): pTz3.Text = Val(t(4, 5))
pV01p.Text = Val(t(5, 3)): pV01y.Text = Val(t(5, 3)): pV011.Text = Val(t(5, 3)): pV012.Text = Val(t(5, 3)): pV013.Text = Val(t(5, 3)): pV014.Text = Val(t(5, 3))
pV02p.Text = Val(t(5, 4)): pV02y.Text = Val(t(5, 4)): pV021.Text = Val(t(5, 4)): pV022.Text = Val(t(5, 4)): pV023.Text = Val(t(5, 4)): pV024.Text = Val(t(5, 4))
pV03p.Text = Val(t(5, 5)): pV03Y.Text = Val(t(5, 5)): pV031.Text = Val(t(5, 5)): pV032.Text = Val(t(5, 5)): pV033.Text = Val(t(5, 5)): pV034.Text = Val(t(5, 5))
End Sub

''''''''''''''''''''''''''''''''METEO''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
Dim MyFile
vvod = 0
dmk = pDMK
If dmk = False Then vvod = 1:
hmet = phmet: Tv = pTv: h = pH: Aw = pAw: W = pW
If dmk = False Then Dsnos = pW: Aw = Aw - 100
        If Aw <= 0 Then Aw = Aw + 6000
        pAw02.Text = Aw / 100 + 1: pAw04.Text = Aw / 100 + 2: pAw08.Text = Aw / 100 + 3: pAw12.Text = Aw / 100 + 3: pAw16.Text = Aw / 100 + 4: pAw20.Text = Aw / 100 + 4: pAw24.Text = Aw / 100 + 4: pAw30.Text = Aw / 100 + 5: pAw40.Text = Aw / 100 + 5: pAw50.Text = Aw / 100 + 5: pAw60.Text = Aw / 100 + 5
 If Tv > 0 And Tv <= 15 Then
        Tv = Tv + 1
        ElseIf Tv > 15 And Tv <= 25 Then
        Tv = Tv + 2
        ElseIf Tv > 25 And Tv <= 35 Then
        Tv = Tv + 3
        ElseIf Tv > 35 Then
        Tv = Tv + 4
        Else
 End If
        ddt = Tv - 15
  If ddt >= 0 Then
        t200 = ddt: t400 = ddt: t800 = ddt: t1200 = ddt: t1600 = ddt: t2000 = ddt: t2400 = ddt: t3000 = ddt: t4000 = ddt: GoTo 2153
        Else
  End If
    If ddt <= 0 Then
        MyFile = FreeFile
        Open App.Path & "\T.ccc" For Input As #MyFile
        Else
    End If
    If ddt <= 0 Then ddt = ddt * (-1) + 50
2141 If EOF(1) Then GoTo 2142
 Input #MyFile, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10
 If ddt = t1 Then t200 = t2: t400 = t3: t800 = t4: t1200 = t5: t1600 = t6: t2000 = t7: t2400 = t8: t3000 = t9: t4000 = t10: GoTo 2142
  GoTo 2141
21412
2142   Close #MyFile
2153
 WETER vvod, W, Dsnos, w200, w400, w800, w1200, w1600, w2000, w2400, w3000, w4000
pW02.Text = w200: pW04.Text = w400: pW08.Text = w800: pW12.Text = w1200: pW16.Text = w1600: pW20.Text = w2000: pW24.Text = w2400: pW30.Text = w3000: pW40.Text = w4000: pW50.Text = w4000: pW60.Text = w4000
pdT02.Text = t200: pdT04.Text = t400: pdT08.Text = t800: pdT12.Text = t1200: pdT16.Text = t1600: pdT20.Text = t2000: pdT24.Text = t2400: pdT30.Text = t3000: pdT40.Text = t4000: pdT50.Text = t4000: pDt60.Text = t4000
End Sub
''''''''''''''''''''''''''''WETER'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function WETER(ByVal vvod As Integer, ByVal W As Integer, ByVal Dsnos As Integer, w200, w400, w800, w1200, w1600, w2000, w2400, w3000, w4000) As Single
If vvod = 0 Then
    Open App.Path & "\weter" For Input As #1
    If W = 0 Then w200 = 0: w400 = 0: w800 = 0: w1200 = 0: w1600 = 0: w2000 = 0: w2400 = 0: w3000 = 0: w4000 = 0: GoTo 2155
2154    If EOF(1) Then GoTo 2155
        Input #1, w1, w2, w3, w4, w5, w6, w7, w8, w9, w10
        If W = w1 Then w200 = w2: w400 = w3: w800 = w4: w1200 = w5: w1600 = w6: w2000 = w7: w2400 = w8: w3000 = w9: w4000 = w10: GoTo 2155
        GoTo 2154
2155    Close #1
        Else
        If Dsnos < 30 Then
                Dsnos = 0
                ElseIf Dsnos >= 30 And Dsnos < 40 Then
                Dsnos = 30
                ElseIf Dsnos >= 40 And Dsnos < 50 Then
                Dsnos = 40
                ElseIf Dsnos >= 50 And Dsnos < 60 Then
                Dsnos = 50
                ElseIf Dsnos >= 60 And Dsnos < 70 Then
                Dsnos = 60
                ElseIf Dsnos >= 70 And Dsnos < 80 Then
                Dsnos = 70
                ElseIf Dsnos >= 80 And Dsnos < 90 Then
                Dsnos = 80
                ElseIf Dsnos >= 90 And Dsnos < 100 Then
                Dsnos = 90
                ElseIf Dsnos >= 100 And Dsnos < 110 Then
                Dsnos = 100
                ElseIf Dsnos >= 110 And Dsnos < 120 Then
                Dsnos = 110
                ElseIf Dsnos >= 120 And Dsnos < 130 Then
                Dsnos = 120
                ElseIf Dsnos >= 130 And Dsnos < 140 Then
                Dsnos = 130
                ElseIf Dsnos >= 140 And Dsnos < 150 Then
                Dsnos = 140
                Else
                Dsnos = 150
        End If
        Open App.Path & "\weter.wr2" For Input As #1
21551  If EOF(1) Then GoTo 21552
        Input #1, w1, w2, w3, w4, w5, w6, w7, w8, w9, w10
       If Dsnos = 0 Then w200 = 0: w400 = 0: w800 = 0: w1200 = 0: w1600 = 0: w2000 = 0: w2400 = 0: w3000 = 0: w4000 = 0: GoTo 21552
        If Dsnos = w1 Then
        w200 = w2: w400 = w3: w800 = w4: w1200 = w5: w1600 = w6: w2000 = w7: w2400 = w8: w3000 = w9: w4000 = w10
        GoTo 21552
        Else
        GoTo 21551
        End If
21552   Close #1
End If
End Function

Private Sub Command2_Click()
ZapisZeli.Show
End Sub

Private Sub Command3_Click()
Shell "D:\YO_NA\sas\sasplanet", vbNormalFocus
End Sub

Private Sub Command4_Click()
ZapisBP.Show
End Sub

Private Sub Command5_Click()
BP.Hide
End Sub

Private Sub Command6_Click()
SobKontrol.Show
End Sub


Private Sub Form_Load()
Dim t(1 To 10) As String
Dim i As Integer

941 Open "D:\YO_NA\optabl" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 942
 Input #1, t(1), t(2), t(3), t(4)
pplOP1.AddItem t(1)
Loop
942 Close #1
9412 Open "D:\YO_NA\optabl" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 9422
 Input #1, t(1), t(2), t(3), t(4)
pplOP2.AddItem t(1)
Loop
9422 Close #1
9413 Open "D:\YO_NA\optabl" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 9423
 Input #1, t(1), t(2), t(3), t(4)
pplOP3.AddItem t(1)
Loop
9423 Close #1
94131 Open "D:\YO_NA\knptabl" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 94231
 Input #1, t(1), t(2), t(3), t(4)
pplNP1.AddItem t(1)
Loop
94231 Close #1
94132 Open "D:\YO_NA\knptabl" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 94232
 Input #1, t(1), t(2), t(3), t(4)
pplNP2.AddItem t(1)
Loop
94232 Close #1
94133 Open "D:\YO_NA\knptabl" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 94233
 Input #1, t(1), t(2), t(3), t(4)
pplNP3.AddItem t(1)
Loop
94233 Close #1
94134 Open "D:\YO_NA\knptabl" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 94234
 Input #1, t(1), t(2), t(3), t(4)
pplNP4.AddItem t(1)
Loop
94234 Close #1
94135 Open "D:\YO_NA\knptabl" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 94235
 Input #1, t(1), t(2), t(3), t(4)
pplNP5.AddItem t(1)
Loop
94235 Close #1
End Sub

Private Sub pDMK_Click()
Dim dmk As String
dmk = pDMK
If dmk = False Then
    Label45.Caption = "Дснос="
    Else
    Label45.Caption = "W="
End If
End Sub

Private Sub pdT02_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw02.Text = ""
pAw02.SetFocus
Else
End If
End Sub
Private Sub pAw02_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW02.Text = ""
pW02.SetFocus
Else
End If
End Sub

Private Sub polMeteo_Click()
Dim dtT As Single, Aw As Single, W As Single
 Dim h As Single
Dim t(1 To 10) As String
Open "D:\YO_NA\sostMeteooo\dist\meteo.txt" For Input As #1
1: If EOF(1) Then GoTo 5
Input #1, t(1), t(2), t(3), t(4)
If t(1) < 1000 Then
phmet.Text = t(1): h = t(2)
    If h > 500 Then
        h = 500 - h + 750
    Else
        h = h + 750
    End If
    pH.Text = h
GoTo 5
Else
End If
GoTo 1
5: Close #1
PolychitByl 2, dtT, Aw, W
pdT02.Text = dtT: pAw02.Text = Aw: pW02.Text = W
PolychitByl 4, dtT, Aw, W
pdT04.Text = dtT: pAw04.Text = Aw: pW04.Text = W
PolychitByl 8, dtT, Aw, W
pdT08.Text = dtT: pAw08.Text = Aw: pW08.Text = W
PolychitByl 12, dtT, Aw, W
pdT12.Text = dtT: pAw12.Text = Aw: pW12.Text = W
PolychitByl 16, dtT, Aw, W
pdT16.Text = dtT: pAw16.Text = Aw: pW16.Text = W
PolychitByl 20, dtT, Aw, W
pdT20.Text = dtT: pAw20.Text = Aw: pW20.Text = W
PolychitByl 24, dtT, Aw, W
pdT24.Text = dtT: pAw24.Text = Aw: pW24.Text = W
PolychitByl 30, dtT, Aw, W
pdT30.Text = dtT: pAw30.Text = Aw: pW30.Text = W
PolychitByl 40, dtT, Aw, W
pdT40.Text = dtT: pAw40.Text = Aw: pW40.Text = W
PolychitByl 50, dtT, Aw, W
pdT50.Text = dtT: pAw50.Text = Aw: pW50.Text = W
PolychitByl 60, dtT, Aw, W
pDt60.Text = dtT: pAw60.Text = Aw: pW60.Text = W

End Sub

Private Sub pW02_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT04.Text = ""
pdT04.SetFocus
Else
End If
End Sub
Private Sub pdT04_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw04.Text = ""
pAw04.SetFocus
Else
End If
End Sub
Private Sub pAw04_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW04.Text = ""
pW04.SetFocus
Else
End If
End Sub
Private Sub pW04_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT08.Text = ""
pdT08.SetFocus
Else
End If
End Sub
Private Sub pdT08_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw08.Text = ""
pAw08.SetFocus
Else
End If
End Sub
Private Sub pAw08_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW08.Text = ""
pW08.SetFocus
Else
End If
End Sub
Private Sub pW08_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT12.Text = ""
pdT12.SetFocus
Else
End If
End Sub
Private Sub pdT12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw12.Text = ""
pAw12.SetFocus
Else
End If
End Sub
Private Sub pAw12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW12.Text = ""
pW12.SetFocus
Else
End If
End Sub
Private Sub pW12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT16.Text = ""
pdT16.SetFocus
Else
End If
End Sub
Private Sub pdT16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw16.Text = ""
pAw16.SetFocus
Else
End If
End Sub
Private Sub pAw16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW16.Text = ""
pW16.SetFocus
Else
End If
End Sub
Private Sub pW16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT20.Text = ""
pdT20.SetFocus
Else
End If
End Sub
Private Sub pdT20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw20.Text = ""
pAw20.SetFocus
Else
End If
End Sub
Private Sub pAw20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW20.Text = ""
pW20.SetFocus
Else
End If
End Sub
Private Sub pW20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT24.Text = ""
pdT24.SetFocus
Else
End If
End Sub
Private Sub pdT24_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw24.Text = ""
pAw24.SetFocus
Else
End If
End Sub
Private Sub pAw24_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW24.Text = ""
pW24.SetFocus
Else
End If
End Sub
Private Sub pW24_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT30.Text = ""
pdT30.SetFocus
Else
End If
End Sub
Private Sub pdT30_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw30.Text = ""
pAw30.SetFocus
Else
End If
End Sub
Private Sub pAw30_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW30.Text = ""
pW30.SetFocus
Else
End If
End Sub
Private Sub pW30_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT40.Text = ""
pdT40.SetFocus
Else
End If
End Sub
Private Sub pdT40_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw40.Text = ""
pAw40.SetFocus
Else
End If
End Sub
Private Sub pAw40_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW40.Text = ""
pW40.SetFocus
Else
End If
End Sub
Private Sub pW40_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdT50.Text = ""
pdT50.SetFocus
Else
End If
End Sub
Private Sub pdT50_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw50.Text = ""
pAw50.SetFocus
Else
End If
End Sub
Private Sub pAw50_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW50.Text = ""
pW50.SetFocus
Else
End If
End Sub
Private Sub pW50_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pDt60.Text = ""
pDt60.SetFocus
Else
End If
End Sub
Private Sub pdT60_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw60.Text = ""
pAw60.SetFocus
Else
End If
End Sub
Private Sub pAw60_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW60.Text = ""
pW60.SetFocus
Else
End If
End Sub
Private Sub PosmNP_Click()
VuvodNP.Show
End Sub

Private Sub pplNP1_Click()
Dim t(1 To 10) As String
Dim nkp As String
Dim Xkp As Single, Ykp As Single, hkp As Single
nkp = pplNP1
  Open "D:\YO_NA\knptabl" For Input As #1
4112  If EOF(1) Then GoTo 4111
  Input #1, t(1), t(2), t(3), t(4)
  If t(1) = nkp Then Xkp = t(2): Ykp = t(3): hkp = t(4)
  GoTo 4112
4111 Close
pXkp1.Text = Xkp: pYkp1.Text = Ykp: phkp1.Text = hkp
End Sub

Private Sub pplNP2_Click()
Dim t(1 To 10) As String
Dim nkp As String
Dim Xkp As Single, Ykp As Single, hkp As Single
nkp = pplNP2
  Open "D:\YO_NA\knptabl" For Input As #1
4112  If EOF(1) Then GoTo 4111
  Input #1, t(1), t(2), t(3), t(4)
  If t(1) = nkp Then Xkp = t(2): Ykp = t(3): hkp = t(4)
  GoTo 4112
4111 Close
pXkp2.Text = Xkp: pYkp2.Text = Ykp: phkp2.Text = hkp
End Sub

Private Sub pplNP3_Click()
Dim t(1 To 10) As String
Dim nkp As String
Dim Xkp As Single, Ykp As Single, hkp As Single
nkp = pplNP3
  Open "D:\YO_NA\knptabl" For Input As #1
4112  If EOF(1) Then GoTo 4111
  Input #1, t(1), t(2), t(3), t(4)
  If t(1) = nkp Then Xkp = t(2): Ykp = t(3): hkp = t(4)
  GoTo 4112
4111 Close
pXkp3.Text = Xkp: pYkp3.Text = Ykp: phkp3.Text = hkp
End Sub

Private Sub pplNP4_Click()
Dim t(1 To 10) As String
Dim nkp As String
Dim Xkp As Single, Ykp As Single, hkp As Single
nkp = pplNP4
  Open "D:\YO_NA\knptabl" For Input As #1
4112  If EOF(1) Then GoTo 4111
  Input #1, t(1), t(2), t(3), t(4)
  If t(1) = nkp Then Xkp = t(2): Ykp = t(3): hkp = t(4)
  GoTo 4112
4111 Close
pXkp4.Text = Xkp: pYkp4.Text = Ykp: phkp4.Text = hkp
End Sub

Private Sub pplNP5_Click()
Dim t(1 To 10) As String
Dim nkp As String
Dim Xkp As Single, Ykp As Single, hkp As Single
nkp = pplNP5
  Open "D:\YO_NA\knptabl" For Input As #1
4112  If EOF(1) Then GoTo 4111
  Input #1, t(1), t(2), t(3), t(4)
  If t(1) = nkp Then Xkp = t(2): Ykp = t(3): hkp = t(4)
  GoTo 4112
4111 Close
pXkp5.Text = Xkp: pYkp5.Text = Ykp: phkp5.Text = hkp
End Sub

Private Sub pplOP1_Click()
Dim q(1 To 10) As String
Dim opn As String
Dim Xop As Single, Yop As Single, hop As Single
   Open "D:\YO_NA\optabl" For Input As #1
   opn = pplOP1
21621  If EOF(1) Then GoTo 2163
  Input #1, q(1), q(2), q(3), q(4)
  If opn = q(1) Then Xop = q(2): Yop = q(3): hop = q(4)
  GoTo 21621
2163  Close
pX1.Text = Xop: pY1.Text = Yop: ph1.Text = hop
End Sub

Private Sub pplOP2_Click()
Dim q(1 To 10) As String
Dim opn As String
Dim Xop As Single, Yop As Single, hop As Single
   Open "D:\YO_NA\optabl" For Input As #1
   opn = pplOP2
21621  If EOF(1) Then GoTo 2163
  Input #1, q(1), q(2), q(3), q(4)
  If opn = q(1) Then Xop = q(2): Yop = q(3): hop = q(4)
  GoTo 21621
2163  Close
pX2.Text = Xop: pY2.Text = Yop: ph2.Text = hop
End Sub

Private Sub pplOP3_Click()
Dim q(1 To 10) As String
Dim opn As String
Dim Xop As Single, Yop As Single, hop As Single
   Open "D:\YO_NA\optabl" For Input As #1
   opn = pplOP3
21621  If EOF(1) Then GoTo 2163
  Input #1, q(1), q(2), q(3), q(4)
  If opn = q(1) Then Xop = q(2): Yop = q(3): hop = q(4)
  GoTo 21621
2163  Close
pX3.Text = Xop: pY3.Text = Yop: ph3.Text = hop
End Sub

Private Sub pVR2_Click()
Dim vr2 As String
vr2 = pVR2
If vr2 = True Then
    Label45.Caption = "Дснос="
    Else
    Label45.Caption = "W="
End If
End Sub
Private Sub pX1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pY1.Text = ""
pY1.SetFocus
End If
End Sub
Private Sub pY1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ph1.Text = ""
ph1.SetFocus
End If
End Sub
Private Sub ph1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pOH1.Text = ""
pOH1.SetFocus
End If
End Sub

Private Sub pX2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pY2.Text = ""
pY2.SetFocus
End If
End Sub
Private Sub pY2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ph2.Text = ""
ph2.SetFocus
End If
End Sub
Private Sub ph2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pOH2.Text = ""
pOH2.SetFocus
End If
End Sub
Private Sub pX3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pY3.Text = ""
pY3.SetFocus
End If
End Sub
Private Sub pY3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ph3.Text = ""
ph3.SetFocus
End If
End Sub
Private Sub ph3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pOH3.Text = ""
pOH3.SetFocus
End If
End Sub
Private Sub pXkp1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYkp1.Text = ""
pYkp1.SetFocus
End If
End Sub
Private Sub pYkp1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phkp1.Text = ""
phkp1.SetFocus
End If
End Sub
Private Sub phkp1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pXkp2.Text = ""
pXkp2.SetFocus
End If
End Sub

Private Sub pXkp2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYkp2.Text = ""
pYkp2.SetFocus
End If
End Sub
Private Sub pYkp2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phkp2.Text = ""
phkp2.SetFocus
End If
End Sub
Private Sub pXkp3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYkp3.Text = ""
pYkp3.SetFocus
End If
End Sub
Private Sub pYkp3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phkp3.Text = ""
phkp3.SetFocus
End If
End Sub
Private Sub pXkp4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYkp4.Text = ""
pYkp4.SetFocus
End If
End Sub
Private Sub pYkp4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phkp4.Text = ""
phkp4.SetFocus
End If
End Sub
Private Sub pXkp5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYkp5.Text = ""
pYkp5.SetFocus
End If
End Sub
Private Sub pYkp5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phkp5.Text = ""
phkp5.SetFocus
End If
End Sub
Private Sub pV01P_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV01y.Text = ""
pV01y.SetFocus
End If
End Sub
Private Sub pV01Y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV011.Text = ""
pV011.SetFocus
End If
End Sub
Private Sub pV011_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV012.Text = ""
pV012.SetFocus
End If
End Sub
Private Sub pV012_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV013.Text = ""
pV013.SetFocus
End If
End Sub
Private Sub pV013_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV014.Text = ""
pV014.SetFocus
End If
End Sub
Private Sub pV02P_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV02y.Text = ""
pV02y.SetFocus
End If
End Sub
Private Sub pV02Y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV021.Text = ""
pV021.SetFocus
End If
End Sub
Private Sub pV021_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV022.Text = ""
pV022.SetFocus
End If
End Sub
Private Sub pV022_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV023.Text = ""
pV023.SetFocus
End If
End Sub
Private Sub pV023_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV024.Text = ""
pV024.SetFocus
End If
End Sub
Private Sub pV03P_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV03Y.Text = ""
pV03Y.SetFocus
End If
End Sub
Private Sub pV03Y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV031.Text = ""
pV031.SetFocus
End If
End Sub
Private Sub pV031_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV032.Text = ""
pV032.SetFocus
End If
End Sub
Private Sub pV032_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV033.Text = ""
pV033.SetFocus
End If
End Sub
Private Sub pV033_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pV034.Text = ""
pV034.SetFocus
End If
End Sub
Private Sub phmet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pH.Text = ""
pH.SetFocus
End If
End Sub
Private Sub pH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pTv.Text = ""
pTv.SetFocus
End If
End Sub
Private Sub pTv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pAw.Text = ""
pAw.SetFocus
End If
End Sub
Private Sub pAw_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pW.Text = ""
pW.SetFocus
End If
End Sub
Function PolychitByl(ByVal hbyl As Single, dtT As Single, Aw As Single, W As Single) As Single
Dim t(1 To 4) As Single
Open "D:\YO_NA\sostMeteooo\dist\meteo.txt" For Input As #1
11: If EOF(1) Then GoTo 51
Input #1, t(1), t(2), t(3), t(4)
If t(1) = hbyl Then
dtT = t(2): Aw = t(3): W = t(4): GoTo 51
Else
End If
GoTo 11
51: Close #1
End Function
