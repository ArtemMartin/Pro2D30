VERSION 5.00
Begin VB.Form Shest6Oryd 
   Caption         =   "Управление 9-ю орюдиями (типа АСУНО)"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton otprKom 
      BackColor       =   &H00FF8080&
      Caption         =   "Отправить команду"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   17000
      Style           =   1  'Graphical
      TabIndex        =   182
      Top             =   100
      Width           =   2000
   End
   Begin VB.CommandButton bpClick 
      BackColor       =   &H00FF8080&
      Caption         =   "БП"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   14500
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   100
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "КОНТРОЛЬ"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   10000
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Пристрелка dX, dY"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   16600
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   6950
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Пристрелка Х,У"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   14900
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   9250
      Width           =   1500
   End
   Begin VB.CommandButton prpoNZR 
      BackColor       =   &H0080FF80&
      Caption         =   "Пристрелка по НЗР"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   14900
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   8100
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Цель для каждого"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   18500
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   6950
      Width           =   1500
   End
   Begin VB.CommandButton PrisDAK 
      BackColor       =   &H0080FF80&
      Caption         =   "Пристрелка ДАК"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   14900
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   6950
      Width           =   1500
   End
   Begin VB.CommandButton Vuxod 
      BackColor       =   &H008080FF&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   18720
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   9600
      Width           =   1200
   End
   Begin VB.Frame Frame4 
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
      Height          =   3700
      Left            =   14500
      TabIndex        =   94
      Top             =   1500
      Width           =   4815
      Begin VB.ComboBox nPlZeli 
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
         Left            =   2300
         TabIndex        =   131
         Text            =   "0"
         Top             =   2800
         Width           =   2200
      End
      Begin VB.CommandButton OZ6Oryd 
         BackColor       =   &H00FF8080&
         Caption         =   "Решить"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   2700
         Width           =   1400
      End
      Begin VB.TextBox pGlc 
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
         Left            =   2900
         TabIndex        =   104
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pFrc 
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
         Left            =   2900
         TabIndex        =   103
         Text            =   "0"
         Top             =   400
         Width           =   1000
      End
      Begin VB.TextBox phc 
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
         Left            =   600
         TabIndex        =   100
         Text            =   "0"
         Top             =   2000
         Width           =   1500
      End
      Begin VB.TextBox pYc 
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
         Left            =   600
         TabIndex        =   99
         Text            =   "0"
         Top             =   1200
         Width           =   1500
      End
      Begin VB.TextBox pXc 
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
         Left            =   600
         TabIndex        =   98
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.Label Label70 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ План Ц."
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2300
         TabIndex        =   106
         Top             =   2000
         Width           =   1605
      End
      Begin VB.Label Label69 
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
         Left            =   2300
         TabIndex        =   102
         Top             =   1200
         Width           =   500
      End
      Begin VB.Label Label68 
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
         Left            =   2300
         TabIndex        =   101
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label67 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
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
         TabIndex        =   97
         Top             =   2000
         Width           =   500
      End
      Begin VB.Label Label66 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
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
         TabIndex        =   96
         Top             =   1200
         Width           =   500
      End
      Begin VB.Label Label65 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Х="
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
         TabIndex        =   95
         Top             =   400
         Width           =   500
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Установки"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   100
      TabIndex        =   10
      Top             =   100
      Width           =   14100
      Begin VB.ComboBox pRep9 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H00800080&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":0000
         Left            =   12700
         List            =   "Shest6Oryd.frx":000A
         TabIndex        =   191
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.ComboBox pRep8 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":0025
         Left            =   11300
         List            =   "Shest6Oryd.frx":002F
         TabIndex        =   190
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.ComboBox pRep7 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H00404080&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":004A
         Left            =   9900
         List            =   "Shest6Oryd.frx":0054
         TabIndex        =   189
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.ComboBox pRep6 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H00008000&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":006F
         Left            =   8500
         List            =   "Shest6Oryd.frx":0079
         TabIndex        =   188
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.ComboBox pRep5 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H00008080&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":0094
         Left            =   7100
         List            =   "Shest6Oryd.frx":009E
         TabIndex        =   187
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.ComboBox pRep4 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H000040C0&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":00B9
         Left            =   5700
         List            =   "Shest6Oryd.frx":00C3
         TabIndex        =   186
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.ComboBox pRep3 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H00808000&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":00DE
         Left            =   4300
         List            =   "Shest6Oryd.frx":00E8
         TabIndex        =   185
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.ComboBox pRep2 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":0103
         Left            =   2900
         List            =   "Shest6Oryd.frx":010D
         TabIndex        =   184
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.ComboBox pRep1 
         BackColor       =   &H0080FF80&
         Height          =   315
         ItemData        =   "Shest6Oryd.frx":0128
         Left            =   1500
         List            =   "Shest6Oryd.frx":0132
         TabIndex        =   183
         Text            =   "Полная"
         Top             =   6300
         Width           =   1000
      End
      Begin VB.TextBox pvdDov9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   480
         Left            =   12700
         TabIndex        =   168
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   480
         Left            =   12700
         TabIndex        =   167
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYr9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   480
         Left            =   12700
         TabIndex        =   166
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvDovt9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   480
         Left            =   12700
         TabIndex        =   165
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYgt9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   480
         Left            =   12700
         TabIndex        =   164
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   480
         Left            =   12700
         TabIndex        =   163
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   480
         Left            =   12700
         TabIndex        =   162
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvts9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   480
         Left            =   12700
         TabIndex        =   161
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvDov9 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   12700
         TabIndex        =   160
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvN9 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   12700
         TabIndex        =   159
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric9 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   12700
         TabIndex        =   158
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.TextBox pvdDov6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   157
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   156
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYr6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   155
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvDovt6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   154
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYgt6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   153
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   152
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   151
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvts6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   150
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvDov6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   149
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvN6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   148
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric6 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   8500
         TabIndex        =   147
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.TextBox pvdDov3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   4300
         TabIndex        =   145
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   4300
         TabIndex        =   144
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYr3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   4300
         TabIndex        =   143
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvDovt3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   4300
         TabIndex        =   142
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYgt3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   4300
         TabIndex        =   141
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   4300
         TabIndex        =   140
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   4300
         TabIndex        =   139
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvts3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   4300
         TabIndex        =   138
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvDov3 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   4300
         TabIndex        =   137
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvN3 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   4300
         TabIndex        =   136
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric3 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   4300
         TabIndex        =   134
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.TextBox pvdDov8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   93
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   92
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYr8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   91
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvDovt8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   90
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYgt8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   89
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   88
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   87
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvts8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   86
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvDov8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   85
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvN8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   84
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   11300
         TabIndex        =   83
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.TextBox pvdDov7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   9900
         TabIndex        =   82
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   9900
         TabIndex        =   81
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvDovt7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   9900
         TabIndex        =   80
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYr7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   9900
         TabIndex        =   79
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvYgt7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   9900
         TabIndex        =   78
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   9900
         TabIndex        =   77
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   9900
         TabIndex        =   76
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvDov7 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   9900
         TabIndex        =   75
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvts7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   9900
         TabIndex        =   74
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvN7 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   9900
         TabIndex        =   73
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric7 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   9900
         TabIndex        =   72
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.TextBox pvdDov5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   7100
         TabIndex        =   71
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   7100
         TabIndex        =   70
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYr5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   7100
         TabIndex        =   69
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvDovt5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   7100
         TabIndex        =   68
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYgt5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   7100
         TabIndex        =   67
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   7100
         TabIndex        =   66
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   7100
         TabIndex        =   65
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvts5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   7100
         TabIndex        =   64
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvDov5 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   7100
         TabIndex        =   63
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvN5 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   7100
         TabIndex        =   62
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric5 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   7100
         TabIndex        =   61
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.TextBox pvdDov4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   60
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   59
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYr4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   58
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvDovt4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   57
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYgt4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   56
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   55
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   54
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvts4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   53
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvDov4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   52
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvN4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   51
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5700
         TabIndex        =   50
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.TextBox pvdDov2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   49
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   48
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYr2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   47
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvDovt2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   46
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYgt2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   45
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   44
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   43
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvts2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   42
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvDov2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   41
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvN2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   40
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   480
         Left            =   2900
         TabIndex        =   39
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.TextBox pvdDov1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1500
         TabIndex        =   38
         Text            =   "0"
         Top             =   5700
         Width           =   1000
      End
      Begin VB.TextBox pvdD1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1500
         TabIndex        =   37
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYr1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1500
         TabIndex        =   36
         Text            =   "0"
         Top             =   4700
         Width           =   1000
      End
      Begin VB.TextBox pvDovt1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1500
         TabIndex        =   35
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pvYgt1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1500
         TabIndex        =   34
         Text            =   "0"
         Top             =   3700
         Width           =   1000
      End
      Begin VB.TextBox pvDt1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1500
         TabIndex        =   33
         Text            =   "0"
         Top             =   3200
         Width           =   1000
      End
      Begin VB.TextBox pvdXtus1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1500
         TabIndex        =   32
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pvts1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1500
         TabIndex        =   31
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvDov1 
         BackColor       =   &H00C0C0FF&
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
         Left            =   1500
         TabIndex        =   30
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvN1 
         BackColor       =   &H00C0C0FF&
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
         Left            =   1500
         TabIndex        =   29
         Text            =   "0"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox pvPric1 
         BackColor       =   &H00C0C0FF&
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
         Left            =   1500
         TabIndex        =   28
         Text            =   "0"
         Top             =   700
         Width           =   1000
      End
      Begin VB.Label labOr9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Кор3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   12700
         TabIndex        =   169
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label labOr6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Сам3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   8500
         TabIndex        =   146
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label labOr3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дес3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   400
         Left            =   4300
         TabIndex        =   135
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label labOr8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Кор2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   400
         Left            =   11300
         TabIndex        =   27
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label labOr7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Кор1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   400
         Left            =   9900
         TabIndex        =   26
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label labOr5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Сам2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   7100
         TabIndex        =   25
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label labOr4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Сам1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   400
         Left            =   5700
         TabIndex        =   24
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label labOr2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дес2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   400
         Left            =   2900
         TabIndex        =   23
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label labOr1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дес1"
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
         Left            =   1500
         TabIndex        =   22
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label Label58 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dДов"
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
         TabIndex        =   21
         Top             =   5700
         Width           =   1000
      End
      Begin VB.Label Label57 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД"
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
         TabIndex        =   20
         Top             =   5200
         Width           =   1000
      End
      Begin VB.Label Label56 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ур"
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
         TabIndex        =   19
         Top             =   4700
         Width           =   1000
      End
      Begin VB.Label Label55 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дов.т"
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
         TabIndex        =   18
         Top             =   4200
         Width           =   1000
      End
      Begin VB.Label Label54 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Уг.т"
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
         TabIndex        =   17
         Top             =   3700
         Width           =   1000
      End
      Begin VB.Label Label53 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дт"
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
         TabIndex        =   16
         Top             =   3200
         Width           =   1000
      End
      Begin VB.Label Label52 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dХтыс"
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
         TabIndex        =   15
         Top             =   2700
         Width           =   1000
      End
      Begin VB.Label Label51 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Полет"
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
         TabIndex        =   14
         Top             =   2200
         Width           =   1000
      End
      Begin VB.Label Label50 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дов.ОН"
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
         Left            =   100
         TabIndex        =   13
         Top             =   1700
         Width           =   1300
      End
      Begin VB.Label Label49 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Труб"
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
         Left            =   100
         TabIndex        =   12
         Top             =   1200
         Width           =   1000
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Прицел"
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
         Left            =   100
         TabIndex        =   11
         Top             =   700
         Width           =   1300
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Заряд Взрыватель Снаряд"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   7000
      Width           =   14100
      Begin VB.ComboBox pOsk3Vzr 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":014D
         Left            =   12200
         List            =   "Shest6Oryd.frx":0160
         TabIndex        =   181
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pOsk3Snar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":0183
         Left            =   12200
         List            =   "Shest6Oryd.frx":0199
         TabIndex        =   180
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pOsk3Zar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":01BA
         Left            =   12200
         List            =   "Shest6Oryd.frx":01D0
         TabIndex        =   179
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox pKal3Snar 
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
         ItemData        =   "Shest6Oryd.frx":01FB
         Left            =   8300
         List            =   "Shest6Oryd.frx":0211
         TabIndex        =   177
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pKal3Vzr 
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
         ItemData        =   "Shest6Oryd.frx":0232
         Left            =   8300
         List            =   "Shest6Oryd.frx":0245
         TabIndex        =   176
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pKal3Zar 
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
         ItemData        =   "Shest6Oryd.frx":0268
         Left            =   8300
         List            =   "Shest6Oryd.frx":027E
         TabIndex        =   175
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox pAks3Snar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":02A9
         Left            =   4400
         List            =   "Shest6Oryd.frx":02BF
         TabIndex        =   173
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pAks3Vzr 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":02E0
         Left            =   4400
         List            =   "Shest6Oryd.frx":02F3
         TabIndex        =   172
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pAks3Zar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":0316
         Left            =   4400
         List            =   "Shest6Oryd.frx":032C
         TabIndex        =   171
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox pOsk2Snar 
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
         ItemData        =   "Shest6Oryd.frx":0357
         Left            =   10900
         List            =   "Shest6Oryd.frx":036D
         TabIndex        =   130
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pOsk1Snar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":038E
         Left            =   9600
         List            =   "Shest6Oryd.frx":03A4
         TabIndex        =   129
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pKal2Snar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":03C5
         Left            =   7000
         List            =   "Shest6Oryd.frx":03DB
         TabIndex        =   128
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pKal1Snar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":03FC
         Left            =   5700
         List            =   "Shest6Oryd.frx":0412
         TabIndex        =   127
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pAks2Snar 
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
         ItemData        =   "Shest6Oryd.frx":0433
         Left            =   3100
         List            =   "Shest6Oryd.frx":0449
         TabIndex        =   126
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pAks1Snar 
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
         ItemData        =   "Shest6Oryd.frx":046A
         Left            =   1800
         List            =   "Shest6Oryd.frx":0480
         TabIndex        =   125
         Text            =   "ОФ"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.ComboBox pOsk2Vzr 
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
         ItemData        =   "Shest6Oryd.frx":04A1
         Left            =   10900
         List            =   "Shest6Oryd.frx":04B4
         TabIndex        =   124
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pOsk1Vzr 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":04D7
         Left            =   9600
         List            =   "Shest6Oryd.frx":04EA
         TabIndex        =   123
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pKal2Vzr 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":050D
         Left            =   7000
         List            =   "Shest6Oryd.frx":0520
         TabIndex        =   122
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pKal1Vzr 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":0543
         Left            =   5700
         List            =   "Shest6Oryd.frx":0556
         TabIndex        =   121
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pAks2Vzr 
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
         ItemData        =   "Shest6Oryd.frx":0579
         Left            =   3100
         List            =   "Shest6Oryd.frx":058C
         TabIndex        =   120
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pAks1Vzr 
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
         ItemData        =   "Shest6Oryd.frx":05AF
         Left            =   1800
         List            =   "Shest6Oryd.frx":05C2
         TabIndex        =   119
         Text            =   "РГМ"
         Top             =   1500
         Width           =   1200
      End
      Begin VB.ComboBox pOsk2Zar 
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
         ItemData        =   "Shest6Oryd.frx":05E5
         Left            =   10900
         List            =   "Shest6Oryd.frx":05FB
         TabIndex        =   118
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox pOsk1Zar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":0626
         Left            =   9600
         List            =   "Shest6Oryd.frx":063C
         TabIndex        =   117
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox pKal2Zar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":0667
         Left            =   7000
         List            =   "Shest6Oryd.frx":067D
         TabIndex        =   116
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox pKal1Zar 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   450
         ItemData        =   "Shest6Oryd.frx":06A8
         Left            =   5700
         List            =   "Shest6Oryd.frx":06BE
         TabIndex        =   115
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox pAks2Zar 
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
         ItemData        =   "Shest6Oryd.frx":06E9
         Left            =   3100
         List            =   "Shest6Oryd.frx":06FF
         TabIndex        =   114
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.ComboBox pAks1Zar 
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
         ItemData        =   "Shest6Oryd.frx":072A
         Left            =   1800
         List            =   "Shest6Oryd.frx":0740
         TabIndex        =   113
         Text            =   "Полн"
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label labeOr9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Кор3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   12200
         TabIndex        =   178
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labeOr6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Сам3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   8300
         TabIndex        =   174
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labeOr3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дес3"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   400
         Left            =   4400
         TabIndex        =   170
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labeOr8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Кор2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   10900
         TabIndex        =   9
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labeOr7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Кор1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   400
         Left            =   9600
         TabIndex        =   8
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labeOr5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Сам2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   7000
         TabIndex        =   7
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labeOr4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Сам1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   400
         Left            =   5700
         TabIndex        =   6
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labeOr2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дес2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   400
         Left            =   3100
         TabIndex        =   5
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label labeOr1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Дес1"
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
         Left            =   1800
         TabIndex        =   4
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Снаряд"
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
         Left            =   100
         TabIndex        =   3
         Top             =   2100
         Width           =   1000
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Взрыватель"
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
         Left            =   100
         TabIndex        =   2
         Top             =   1500
         Width           =   1600
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Заряд"
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
         Left            =   100
         TabIndex        =   1
         Top             =   900
         Width           =   1000
      End
   End
End
Attribute VB_Name = "Shest6Oryd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bpClick_Click()
bp6Oryd.Show
End Sub

Private Sub Command1_Click()
Zelkagdomy6.Show
End Sub

Private Sub Command2_Click()
PrisXYform.Show
End Sub

Private Sub Command3_Click()
PrisdXdYform.Show
End Sub


Private Sub Command4_Click()
KontrolPooryd.Show
End Sub

Private Sub Form_Load()
Dim t(1 To 10) As String
Dim i As Integer
941 Open "D:\YO_NA\zeli" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 942
 Input #1, t(1), t(2), t(3), t(4), t(5), t(6)
nPlZeli.AddItem t(1)
Loop
942 Close #1

Open App.Path & "\BP\bp9.txt" For Input As #1
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr1 = t(1)
labeOr1 = t(1)
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr2 = t(1)
labeOr2 = t(1)
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr3 = t(1)
labeOr3 = t(1)
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr4 = t(1)
labeOr4 = t(1)
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr5 = t(1)
labeOr5 = t(1)
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr6 = t(1)
labeOr6 = t(1)
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr7 = t(1)
labeOr7 = t(1)
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr8 = t(1)
labeOr8 = t(1)
Input #1, t(1), t(2), t(3), t(4), t(5), t(6), t(7)
labOr9 = t(1)
labeOr9 = t(1)
Close #1
End Sub

Private Sub nPlZeli_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = nPlZeli
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Frc = t(4): Glc = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc.Text = Xc: pYc.Text = Yc: phc.Text = hc: pFrc.Text = Frc: pGlc.Text = Glc
End Sub
Private Sub nPlZeli_KeyDown(KeyCode As Integer, Shift As Integer)
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
pXc.Text = Xc: pYc.Text = Yc: phc.Text = hc: pFrc.Text = Frc: pGlc.Text = Glc
    Else
End If
End Sub

Private Sub otprKom_Click()
otprKomPooryd.Show
End Sub

Private Sub OZ6Oryd_Click()
Dim Nop As Integer
Dim Xop As Single, Xc As Single, Yc As Single, hc As Single
Dim rep As String
Dim preps1 As Single, N1 As Single, dovisch1 As Single, ts As Single, dXtus As Single
Dim Dt1 As Single, Ygolt As Single, Dovort1 As Single, yroven As Single, popvd1 As Single
Dim popvnap1 As Single

'записать номер цели в файл
Open App.Path & "\numberZeli" For Output As #1
Write #1, nPlZeli
Close #1

Xc = pXc: Yc = pYc: hc = phc
Xop = bp6Oryd.pAks1X: Nop = 1: rep = pRep1
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric1.Text = preps1: pvN1.Text = Round(N1): pvDov1.Text = dovisch1: pvdXtus1.Text = dXtus: pvts1.Text = ts
    pvDt1.Text = Dt1: pvYgt1.Text = Ygolt: pvDovt1.Text = Dovort1: pvYr1.Text = Round(yroven): pvdD1.Text = Round(popvd1)
    pvdDov1.Text = Round(popvnap1)
Else
    pvPric1.Text = 0: pvN1.Text = 0: pvDov1.Text = 0: pvdXtus1.Text = 0: pvts1.Text = 0
    pvDt1.Text = 0: pvYgt1.Text = 0: pvDovt1.Text = 0: pvYr1.Text = 0: pvdD1.Text = 0
    pvdDov1.Text = 0
End If
Xop = bp6Oryd.pAks2X: Nop = 2: rep = pRep2
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric2.Text = preps1: pvN2.Text = Round(N1): pvDov2.Text = dovisch1: pvdXtus2.Text = dXtus: pvts2.Text = ts
    pvDt2.Text = Dt1: pvYgt2.Text = Ygolt: pvDovt2.Text = Dovort1: pvYr2.Text = Round(yroven): pvdD2.Text = Round(popvd1)
    pvdDov2.Text = Round(popvnap1)
Else
    pvPric2.Text = 0: pvN2.Text = 0: pvDov2.Text = 0: pvdXtus2.Text = 0: pvts2.Text = 0
    pvDt2.Text = 0: pvYgt2.Text = 0: pvDovt2.Text = 0: pvYr2.Text = 0: pvdD2.Text = 0
    pvdDov2.Text = 0
End If
Xop = bp6Oryd.pAks3X: Nop = 3: rep = pRep3
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric3.Text = preps1: pvN3.Text = Round(N1): pvDov3.Text = dovisch1: pvdXtus3.Text = dXtus: pvts3.Text = ts
    pvDt3.Text = Dt1: pvYgt3.Text = Ygolt: pvDovt3.Text = Dovort1: pvYr3.Text = Round(yroven): pvdD3.Text = Round(popvd1)
    pvdDov3.Text = Round(popvnap1)
Else
    pvPric3.Text = 0: pvN3.Text = 0: pvDov3.Text = 0: pvdXtus3.Text = 0: pvts3.Text = 0
    pvDt3.Text = 0: pvYgt3.Text = 0: pvDovt3.Text = 0: pvYr3.Text = 0: pvdD3.Text = 0
    pvdDov3.Text = 0
End If

Xop = bp6Oryd.pKal1X: Nop = 4: rep = pRep4
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric4.Text = preps1: pvN4.Text = Round(N1): pvDov4.Text = dovisch1: pvdXtus4.Text = dXtus: pvts4.Text = ts
    pvDt4.Text = Dt1: pvYgt4.Text = Ygolt: pvDovt4.Text = Dovort1: pvYr4.Text = Round(yroven): pvdD4.Text = Round(popvd1)
    pvdDov4.Text = Round(popvnap1)
Else
    pvPric4.Text = 0: pvN4.Text = 0: pvDov4.Text = 0: pvdXtus4.Text = 0: pvts4.Text = 0
    pvDt4.Text = 0: pvYgt4.Text = 0: pvDovt4.Text = 0: pvYr4.Text = 0: pvdD4.Text = 0
    pvdDov4.Text = 0
End If
Xop = bp6Oryd.pKal2X: Nop = 5: rep = pRep5
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric5.Text = preps1: pvN5.Text = Round(N1): pvDov5.Text = dovisch1: pvdXtus5.Text = dXtus: pvts5.Text = ts
    pvDt5.Text = Dt1: pvYgt5.Text = Ygolt: pvDovt5.Text = Dovort1: pvYr5.Text = Round(yroven): pvdD5.Text = Round(popvd1)
    pvdDov5.Text = Round(popvnap1)
Else
    pvPric5.Text = 0: pvN5.Text = 0: pvDov5.Text = 0: pvdXtus5.Text = 0: pvts5.Text = 0
    pvDt5.Text = 0: pvYgt5.Text = 0: pvDovt5.Text = 0: pvYr5.Text = 0: pvdD5.Text = 0
    pvdDov5.Text = 0
End If
Xop = bp6Oryd.pKal3X: Nop = 6: rep = pRep6
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric6.Text = preps1: pvN6.Text = Round(N1): pvDov6.Text = dovisch1: pvdXtus6.Text = dXtus: pvts6.Text = ts
    pvDt6.Text = Dt1: pvYgt6.Text = Ygolt: pvDovt6.Text = Dovort1: pvYr6.Text = Round(yroven): pvdD6.Text = Round(popvd1)
    pvdDov6.Text = Round(popvnap1)
Else
    pvPric6.Text = 0: pvN6.Text = 0: pvDov6.Text = 0: pvdXtus6.Text = 0: pvts6.Text = 0
    pvDt6.Text = 0: pvYgt6.Text = 0: pvDovt6.Text = 0: pvYr6.Text = 0: pvdD6.Text = 0
    pvdDov6.Text = 0
End If

Xop = bp6Oryd.pOsk1X: Nop = 7: rep = pRep7
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric7.Text = preps1: pvN7.Text = Round(N1): pvDov7.Text = dovisch1: pvdXtus7.Text = dXtus: pvts7.Text = ts
    pvDt7.Text = Dt1: pvYgt7.Text = Ygolt: pvDovt7.Text = Dovort1: pvYr7.Text = Round(yroven): pvdD7.Text = Round(popvd1)
    pvdDov7.Text = Round(popvnap1)
Else
    pvPric7.Text = 0: pvN7.Text = 0: pvDov7.Text = 0: pvdXtus7.Text = 0: pvts7.Text = 0
    pvDt7.Text = 0: pvYgt7.Text = 0: pvDovt7.Text = 0: pvYr7.Text = 0: pvdD7.Text = 0
    pvdDov7.Text = 0
End If
Xop = bp6Oryd.pOsk2X: Nop = 8: rep = pRep8
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric8.Text = preps1: pvN8.Text = Round(N1): pvDov8.Text = dovisch1: pvdXtus8.Text = dXtus: pvts8.Text = ts
    pvDt8.Text = Dt1: pvYgt8.Text = Ygolt: pvDovt8.Text = Dovort1: pvYr8.Text = Round(yroven): pvdD8.Text = Round(popvd1)
    pvdDov8.Text = Round(popvnap1)
Else
    pvPric8.Text = 0: pvN8.Text = 0: pvDov8.Text = 0: pvdXtus8.Text = 0: pvts8.Text = 0
    pvDt8.Text = 0: pvYgt8.Text = 0: pvDovt8.Text = 0: pvYr8.Text = 0: pvdD8.Text = 0
    pvdDov8.Text = 0
End If
Xop = bp6Oryd.pOsk3X: Nop = 9: rep = pRep9
If Xop <> 0 Then
    OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    pvPric9.Text = preps1: pvN9.Text = Round(N1): pvDov9.Text = dovisch1: pvdXtus9.Text = dXtus: pvts9.Text = ts
    pvDt9.Text = Dt1: pvYgt9.Text = Ygolt: pvDovt9.Text = Dovort1: pvYr9.Text = Round(yroven): pvdD9.Text = Round(popvd1)
    pvdDov9.Text = Round(popvnap1)
Else
    pvPric9.Text = 0: pvN9.Text = 0: pvDov9.Text = 0: pvdXtus9.Text = 0: pvts9.Text = 0
    pvDt9.Text = 0: pvYgt9.Text = 0: pvDovt9.Text = 0: pvYr9.Text = 0: pvdD9.Text = 0
    pvdDov9.Text = 0
End If

End Sub

Sub OZ6or(ByVal Nop As Single, ByVal Xc As Single, ByVal Yc As Single, ByVal hc As Single, ByVal rep As String, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1)
Dim v01 As Single, pop_v_N As Single, pop_v_Nk As Single, dN As Single, dddt1 As Single
Dim epsDnO13 As Single, Xop As Single, Yop As Single, hop As Single, tz As Single
Dim OH As Single
Dim zar As String, snar As String, vzriv As String
If Nop = 1 Then
    Xop = bp6Oryd.pAks1X: Yop = bp6Oryd.pAks1Y: hop = bp6Oryd.pAks1h
    tz = bp6Oryd.pAks1Tz: OH = bp6Oryd.pAks1ON
    zar = pAks1Zar: snar = pAks1Snar: vzriv = pAks1Vzr: v01 = bp6Oryd.pAks1V0
    ElseIf Nop = 2 Then
    Xop = bp6Oryd.pAks2X: Yop = bp6Oryd.pAks2Y: hop = bp6Oryd.pAks2h: tz = bp6Oryd.pAks2Tz: OH = bp6Oryd.pAks2ON
    zar = pAks2Zar: snar = pAks2Snar: vzriv = pAks2Vzr: v01 = bp6Oryd.pAks2V0
    ElseIf Nop = 3 Then
    Xop = bp6Oryd.pAks3X: Yop = bp6Oryd.pAks3Y: hop = bp6Oryd.pAks3h: tz = bp6Oryd.pAks3Tz: OH = bp6Oryd.pAks3ON
    zar = pAks3Zar: snar = pAks3Snar: vzriv = pAks3Vzr: v01 = bp6Oryd.pAks3V0
    ElseIf Nop = 4 Then
    Xop = bp6Oryd.pKal1X: Yop = bp6Oryd.pKal1Y: hop = bp6Oryd.pKal1h: tz = bp6Oryd.pKal1Tz: OH = bp6Oryd.pKal1ON
    zar = pKal1Zar: snar = pKal1Snar: vzriv = pKal1Vzr: v01 = bp6Oryd.pKal1V0
    ElseIf Nop = 5 Then
    Xop = bp6Oryd.pKal2X: Yop = bp6Oryd.pKal2Y: hop = bp6Oryd.pKal2h: tz = bp6Oryd.pKal2Tz: OH = bp6Oryd.pKal2ON
    zar = pKal2Zar: snar = pKal2Snar: vzriv = pKal2Vzr: v01 = bp6Oryd.pKal2V0
    ElseIf Nop = 6 Then
    Xop = bp6Oryd.pKal3X: Yop = bp6Oryd.pKal3Y: hop = bp6Oryd.pKal3h: tz = bp6Oryd.pKal3Tz: OH = bp6Oryd.pKal3ON
    zar = pKal3Zar: snar = pKal3Snar: vzriv = pKal3Vzr: v01 = bp6Oryd.pKal3V0
    ElseIf Nop = 7 Then
    Xop = bp6Oryd.pOsk1X: Yop = bp6Oryd.pOsk1Y: hop = bp6Oryd.pOsk1h: tz = bp6Oryd.pOsk1Tz: OH = bp6Oryd.pOsk1ON
    zar = pOsk1Zar: snar = pOsk1Snar: vzriv = pOsk1Vzr: v01 = bp6Oryd.pOsk1V0
    ElseIf Nop = 8 Then
    Xop = bp6Oryd.pOsk2X: Yop = bp6Oryd.pOsk2Y: hop = bp6Oryd.pOsk2h: tz = bp6Oryd.pOsk2Tz: OH = bp6Oryd.pOsk2ON
    zar = pOsk2Zar: snar = pOsk2Snar: vzriv = pOsk2Vzr: v01 = bp6Oryd.pOsk2V0
    Else
    Xop = bp6Oryd.pOsk3X: Yop = bp6Oryd.pOsk3Y: hop = bp6Oryd.pOsk3h: tz = bp6Oryd.pOsk3Tz: OH = bp6Oryd.pOsk3ON
    zar = pOsk3Zar: snar = pOsk3Snar: vzriv = pOsk3Vzr: v01 = bp6Oryd.pOsk3V0
End If
      '1B
50:
Dim ras As Single, h As Single, hmet As Single, Xc1 As Single, Yc1 As Single
Dim hc1 As Single, dx1 As Single, dy1 As Single, dh1 As Single
Dim Pi As Single, dhh1 As Single

ras = 0: h = BP.ph:  hmet = BP.phmet
If h = 0 Then h = 750
dhh1 = (h - 750) + ((hmet - hop) / 10)
   Xc1 = Xc: Yc1 = Yc: hc1 = hc
   dx1 = Xc - Xop
60: dy1 = Yc - Yop
61: dh1 = hc - hop
Pi = 3.14159265358
9010: Dt1 = Int(Sqr(dx1 ^ 2 + dy1 ^ 2) + 0.001)

Dim Yr1 As Single, A1 As Single, Ygolt1 As Single, OH1 As Single, Dt As Single
Dim dh As Single, ybyl As Single, zc As Single, dZwc As Single, dXwc As Single
Dim stre As String
Dim dXhc As Single, dXtc As Single, dXv0c As Single, met As Single, ybylc As Single
Dim Wx As Single, Wz As Single, dddt As Single, Ygvozv As Single, Ygpad As Single
Dim Vustra As Single, Vd As Single, tsk As Single, dXtusk As Single, Ygvozvk As Single
Dim Vustrak As Single, Ygpadk As Single, Vdk As Single, dv0 As Single, rep1 As String
Dim dDov1 As Single, Dret1 As Single, dDr1 As Single, popvnap As Single, popvD As Single
Dim Dtk As Single, popvnapk As Single, popvdk As Single
Dim Dtisch As Single, Dtischk As Single, kPop As Single, Disch As Single, Disch1 As Single
Dim Kpopnap As Single, dhh As Single, dV00 As Single, Pricisch As Single, N As Single
Dim vsem As Single, dNtus As Single, vrv As Single, Pric1 As Single, Yr As Single
Dim Yrr As Single, dNtus1 As Single, Ygpad1 As Single, Ygvozv1 As Single, Vustra1 As Single
Dim ts1 As Single, dXtus11 As Single, kpe As Single, daep As Single

9110: Yr1 = CInt((dh1 / (Dt1 * 0.001 + 0.001)) * 0.95)
100: A1 = Abs(Atn(dy1 / (dx1 + 0.001)) / Pi * 30) * 100
101: If dx1 > 0 And dy1 > 0 Then Ygolt1 = Round(A1)
102: If dx1 < 0 And dy1 > 0 Then Ygolt1 = Round(3000 - A1)
103: If dx1 < 0 And dy1 < 0 Then Ygolt1 = Round(3000 + A1)
104: If dx1 > 0 And dy1 < 0 Then Ygolt1 = Round(6000 - A1)
10411: If Ygolt1 <= 1500 And OH >= 4500 Then
      Dovort1 = Ygolt1 + 6000 - OH
      ElseIf OH <= 1500 And Ygolt1 >= 4500 Then
      Dovort1 = Ygolt1 - (OH + 6000)
      Else
      Dovort1 = Ygolt1 - OH
      End If
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1
Dim pozuvnoiOP As String

pozuvnoiOP = getPozuvnOP(Nop)
OZ.msgVelikaDalnost snar, zar, pozuvnoiOP, Dt

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
       dddt1 = dddt
       
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       OZ.poddV0 tz, zar, dv0
       rep1 = pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1
       dDr1 = REPER.pvdD1: dN = REPER.pvdN1
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
        Xc1 = Xc: Yc1 = Yc: hc1 = hc: yroven = Yr1
End Sub


Private Sub pRep1_Click()
If pRep1 = "Полная" Then
    pRep1.BackColor = &H80FF80
    Else
    pRep1.BackColor = &H8080FF
End If
End Sub
Private Sub pRep2_Click()
If pRep2 = "Полная" Then
    pRep2.BackColor = &H80FF80
    Else
    pRep2.BackColor = &H8080FF
End If
End Sub
Private Sub pRep3_Click()
If pRep3 = "Полная" Then
    pRep3.BackColor = &H80FF80
    Else
    pRep3.BackColor = &H8080FF
End If
End Sub
Private Sub pRep4_Click()
If pRep4 = "Полная" Then
    pRep4.BackColor = &H80FF80
    Else
    pRep4.BackColor = &H8080FF
End If
End Sub
Private Sub pRep5_Click()
If pRep5 = "Полная" Then
    pRep5.BackColor = &H80FF80
    Else
    pRep5.BackColor = &H8080FF
End If
End Sub
Private Sub pRep6_Click()
If pRep6 = "Полная" Then
    pRep6.BackColor = &H80FF80
    Else
    pRep6.BackColor = &H8080FF
End If
End Sub
Private Sub pRep7_Click()
If pRep7 = "Полная" Then
    pRep7.BackColor = &H80FF80
    Else
    pRep7.BackColor = &H8080FF
End If
End Sub
Private Sub pRep8_Click()
If pRep8 = "Полная" Then
    pRep8.BackColor = &H80FF80
    Else
    pRep8.BackColor = &H8080FF
End If
End Sub
Private Sub pRep9_Click()
If pRep9 = "Полная" Then
    pRep9.BackColor = &H80FF80
    Else
    pRep9.BackColor = &H8080FF
End If
End Sub

Private Sub PrisDAK_Click()
PrisDAKform.Show
End Sub

Private Sub prpoNZR_Click()
prpoNZRfrm.Show
End Sub

Private Sub pXc_Click()
pXc.Text = ""
End Sub

Private Sub pXc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pYc.Text = ""
    pYc.SetFocus
    Else
End If
End Sub
Private Sub pyc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    phc.Text = ""
    phc.SetFocus
    Else
End If
End Sub

Private Sub Vuxod_Click()
Shest6Oryd.Hide
End Sub

Private Sub vzOpPooryd_Click()
Dim t(5, 8) As String
Dim myArr() As String
Dim i As Integer, j As Integer
Dim stroka As String

'записываем файл в масив
Open "D:\YO_NA\BP_Brig\gsadnPooryd" For Input As #1
10: If EOF(1) Then GoTo 20
  Line Input #1, stroka
   myArr = Split(stroka, ",")
  For i = 0 To 8
        t(j, i) = myArr(i)
        Next i
        j = j + 1
   GoTo 10
20: Close #1
'пердаем данные с масива в поля батарей
bp6Oryd.pAks1X.Text = Val(t(0, 0)): bp6Oryd.pAks1Y.Text = Val(t(1, 0)): bp6Oryd.pAks1h.Text = Val(t(2, 0))
bp6Oryd.pAks1ON.Text = Val(t(3, 0)): bp6Oryd.pAks1Tz.Text = Val(t(4, 0)): bp6Oryd.pAks1V0.Text = Val(t(5, 0))
bp6Oryd.pAks2X.Text = Val(t(0, 1)): bp6Oryd.pAks2Y.Text = Val(t(1, 1)): bp6Oryd.pAks2h.Text = Val(t(2, 1))
bp6Oryd.pAks2ON.Text = Val(t(3, 1)): bp6Oryd.pAks2Tz.Text = Val(t(4, 1)): bp6Oryd.pAks2V0.Text = Val(t(5, 1))
bp6Oryd.pAks3X.Text = Val(t(0, 2)): bp6Oryd.pAks3Y.Text = Val(t(1, 2)): bp6Oryd.pAks3h.Text = Val(t(2, 2))
bp6Oryd.pAks3ON.Text = Val(t(3, 2)): bp6Oryd.pAks3Tz.Text = Val(t(4, 2)): bp6Oryd.pAks3V0.Text = Val(t(5, 2))

'bp6Oryd.pKal1X.Text = Val(t(0, 3)): bp6Oryd.pKal1Y.Text = Val(t(1, 3)): bp6Oryd.pKal1h.Text = Val(t(2, 3))
'bp6Oryd.pKal1ON.Text = Val(t(3, 3)): bp6Oryd.pKal1Tz.Text = Val(t(4, 3)): bp6Oryd.pKal1V0.Text = Val(t(5, 3))
'bp6Oryd.pKal2X.Text = Val(t(0, 4)): bp6Oryd.pKal2Y.Text = Val(t(1, 4)): bp6Oryd.pKal2h.Text = Val(t(2, 4))
'bp6Oryd.pKal2ON.Text = Val(t(3, 4)): bp6Oryd.pKal2Tz.Text = Val(t(4, 4)): bp6Oryd.pKal2V0.Text = Val(t(5, 4))
'bp6Oryd.pKal3X.Text = Val(t(0, 5)): bp6Oryd.pKal3Y.Text = Val(t(1, 5)): bp6Oryd.pKal3h.Text = Val(t(2, 5))
'bp6Oryd.pKal3ON.Text = Val(t(3, 5)): bp6Oryd.pKal3Tz.Text = Val(t(4, 5)): bp6Oryd.pKal3V0.Text = Val(t(5, 5))

'bp6Oryd.pOsk1X.Text = Val(t(0, 6)): bp6Oryd.pOsk1Y.Text = Val(t(1, 6)): bp6Oryd.pOsk1h.Text = Val(t(2, 6))
'bp6Oryd.pOsk1ON.Text = Val(t(3, 6)): bp6Oryd.pOsk1Tz.Text = Val(t(4, 6)): bp6Oryd.pOsk1V0.Text = Val(t(5, 6))
'bp6Oryd.pOsk2X.Text = Val(t(0, 7)): bp6Oryd.pOsk2Y.Text = Val(t(1, 7)): bp6Oryd.pOsk2h.Text = Val(t(2, 7))
'bp6Oryd.pOsk2ON.Text = Val(t(3, 7)): bp6Oryd.pOsk2Tz.Text = Val(t(4, 7)): bp6Oryd.pOsk2V0.Text = Val(t(5, 7))
'bp6Oryd.pOsk3X.Text = Val(t(0, 8)): bp6Oryd.pOsk3Y.Text = Val(t(1, 8)): bp6Oryd.pOsk3h.Text = Val(t(2, 8))
'bp6Oryd.pOsk3ON.Text = Val(t(3, 8)): bp6Oryd.pOsk3Tz.Text = Val(t(4, 8)): bp6Oryd.pOsk3V0.Text = Val(t(5, 8))

End Sub

Sub polPoprPoRepery(ByVal Nop As Integer, ByVal Dt As Single, dD, dDov)
Dim dDrep As Single, dDovRep As Single, dalRtopo As Single
Dim kPop As Single

If Nop = 1 Then
    dalRtopo = ReperPoorydFrm.pvDr1: dDrep = ReperPoorydFrm.pvdD1: dDovRep = ReperPoorydFrm.pvdDov1
    ElseIf Nop = 2 Then
    dalRtopo = ReperPoorydFrm.pvDr2: dDrep = ReperPoorydFrm.pvdD2: dDovRep = ReperPoorydFrm.pvdDov2
    ElseIf Nop = 3 Then
    dalRtopo = ReperPoorydFrm.pvDr3: dDrep = ReperPoorydFrm.pvdD3: dDovRep = ReperPoorydFrm.pvdDov3
    ElseIf Nop = 4 Then
    dalRtopo = ReperPoorydFrm.pvDr4: dDrep = ReperPoorydFrm.pvdD4: dDovRep = ReperPoorydFrm.pvdDov4
    ElseIf Nop = 5 Then
    dalRtopo = ReperPoorydFrm.pvDr5: dDrep = ReperPoorydFrm.pvdD5: dDovRep = ReperPoorydFrm.pvdDov5
    ElseIf Nop = 6 Then
    dalRtopo = ReperPoorydFrm.pvDr6: dDrep = ReperPoorydFrm.pvdD6: dDovRep = ReperPoorydFrm.pvdDov6
    ElseIf Nop = 7 Then
    dalRtopo = ReperPoorydFrm.pvDr7: dDrep = ReperPoorydFrm.pvdD7: dDovRep = ReperPoorydFrm.pvdDov7
    ElseIf Nop = 8 Then
    dalRtopo = ReperPoorydFrm.pvDr8: dDrep = ReperPoorydFrm.pvdD8: dDovRep = ReperPoorydFrm.pvdDov8
    ElseIf Nop = 9 Then
    dalRtopo = ReperPoorydFrm.pvDr9: dDrep = ReperPoorydFrm.pvdD9: dDovRep = ReperPoorydFrm.pvdDov9
    Else
End If

dD = (dDrep / (dalRtopo + 0.0001)) * Dt
dDov = (dDovRep / (dalRtopo + 0.0001)) * Dt

End Sub

Function getPozuvnOP(ByVal Nop As Integer) As String

Select Case Nop
    Case 1
        getPozuvnOP = labOr1
    Case 2
        getPozuvnOP = labOr2
    Case 3
        getPozuvnOP = labOr3
    Case 4
        getPozuvnOP = labOr4
    Case 5
        getPozuvnOP = labOr5
    Case 6
        getPozuvnOP = labOr6
    Case 7
        getPozuvnOP = labOr7
    Case 8
        getPozuvnOP = labOr8
    Case 9
        getPozuvnOP = labOr9
    Case Else
        getPozuvnOP = labOr1
End Select

End Function
