VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.OCX"
Begin VB.Form YO 
   BackColor       =   &H0000C0C0&
   Caption         =   "Управление Огнем"
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
   LinkTopic       =   "YO"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab YO 
      Height          =   11200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20490
      _ExtentX        =   36142
      _ExtentY        =   19764
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   706
      BackColor       =   32768
      ForeColor       =   128
      TabCaption(0)   =   "Боевой порядок"
      TabPicture(0)   =   "Form3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Огневая задача"
      TabPicture(1)   =   "Form3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabPicture(2)   =   "Form3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabPicture(3)   =   "Form3.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabPicture(4)   =   "Form3.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Цель"
         Height          =   9600
         Left            =   -74880
         TabIndex        =   280
         Top             =   600
         Width           =   8535
         Begin VB.TextBox Text108 
            Height          =   400
            Left            =   5500
            TabIndex        =   318
            Text            =   "0"
            Top             =   1400
            Width           =   1000
         End
         Begin VB.TextBox Text107 
            Height          =   400
            Left            =   5500
            TabIndex        =   317
            Text            =   "0"
            Top             =   900
            Width           =   1000
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "решить"
            Height          =   900
            Left            =   6150
            Style           =   1  'Graphical
            TabIndex        =   314
            Top             =   6900
            Width           =   1000
         End
         Begin VB.ComboBox Combo15 
            Height          =   405
            ItemData        =   "Form3.frx":008C
            Left            =   4450
            List            =   "Form3.frx":009F
            TabIndex        =   313
            Text            =   "1"
            Top             =   6400
            Width           =   800
         End
         Begin VB.TextBox Text106 
            Height          =   400
            Left            =   4450
            TabIndex        =   312
            Text            =   "0"
            Top             =   7400
            Width           =   1000
         End
         Begin VB.TextBox Text105 
            Height          =   400
            Left            =   4450
            TabIndex        =   311
            Text            =   "0"
            Top             =   6900
            Width           =   1500
         End
         Begin VB.ComboBox Combo14 
            Height          =   405
            ItemData        =   "Form3.frx":00B2
            Left            =   1000
            List            =   "Form3.frx":00C5
            TabIndex        =   307
            Text            =   "1"
            Top             =   6400
            Width           =   800
         End
         Begin VB.TextBox Text104 
            Height          =   400
            Left            =   1000
            TabIndex        =   306
            Text            =   "0"
            Top             =   7400
            Width           =   1000
         End
         Begin VB.TextBox Text103 
            Height          =   400
            Left            =   1000
            TabIndex        =   305
            Text            =   "0"
            Top             =   6900
            Width           =   1500
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Решить"
            Height          =   900
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   298
            Top             =   3700
            Width           =   1000
         End
         Begin VB.TextBox Text102 
            Height          =   400
            Left            =   1000
            TabIndex        =   297
            Text            =   "0"
            Top             =   4700
            Width           =   1000
         End
         Begin VB.TextBox Text101 
            Height          =   400
            Left            =   1000
            TabIndex        =   296
            Text            =   "0"
            Top             =   4200
            Width           =   1500
         End
         Begin VB.TextBox Text100 
            Height          =   400
            Left            =   1000
            TabIndex        =   295
            Text            =   "0"
            Top             =   3700
            Width           =   1500
         End
         Begin VB.ComboBox Combo13 
            Height          =   405
            ItemData        =   "Form3.frx":00D8
            Left            =   1000
            List            =   "Form3.frx":00EB
            TabIndex        =   294
            Text            =   "1"
            Top             =   3200
            Width           =   800
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Решить"
            Height          =   900
            Left            =   2700
            MaskColor       =   &H00808080&
            Style           =   1  'Graphical
            TabIndex        =   288
            Top             =   900
            Width           =   1000
         End
         Begin VB.TextBox Text99 
            Height          =   405
            Left            =   1000
            TabIndex        =   287
            Text            =   "0"
            Top             =   1900
            Width           =   615
         End
         Begin VB.TextBox Text98 
            Height          =   400
            Left            =   1000
            TabIndex        =   286
            Text            =   "0"
            Top             =   1400
            Width           =   1500
         End
         Begin VB.TextBox Text97 
            Height          =   400
            Left            =   1000
            TabIndex        =   285
            Text            =   "0"
            Top             =   900
            Width           =   1500
         End
         Begin VB.Label Label89 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Фронт"
            Height          =   300
            Left            =   4400
            TabIndex        =   319
            Top             =   900
            Width           =   1000
         End
         Begin VB.Label Label88 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Глубина"
            Height          =   300
            Left            =   4400
            TabIndex        =   316
            Top             =   1400
            Width           =   1000
         End
         Begin VB.Label Label86 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Размеры цели"
            Height          =   300
            Left            =   4400
            TabIndex        =   315
            Top             =   400
            Width           =   2100
         End
         Begin VB.Label Label85 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Мц="
            Height          =   300
            Left            =   3600
            TabIndex        =   310
            Top             =   7400
            Width           =   600
         End
         Begin VB.Label Label84 
            BackColor       =   &H00C0C0C0&
            Caption         =   "А="
            Height          =   300
            Left            =   3600
            TabIndex        =   309
            Top             =   6900
            Width           =   400
         End
         Begin VB.Label Label83 
            BackColor       =   &H00C0C0C0&
            Caption         =   "№ КП="
            Height          =   300
            Left            =   3600
            TabIndex        =   308
            Top             =   6400
            Width           =   900
         End
         Begin VB.Label Label82 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Мц="
            Height          =   300
            Left            =   150
            TabIndex        =   304
            Top             =   7400
            Width           =   600
         End
         Begin VB.Label Label81 
            BackColor       =   &H00C0C0C0&
            Caption         =   "А="
            Height          =   300
            Left            =   150
            TabIndex        =   303
            Top             =   6900
            Width           =   400
         End
         Begin VB.Label Label80 
            BackColor       =   &H00C0C0C0&
            Caption         =   "№ КП="
            Height          =   300
            Left            =   150
            TabIndex        =   302
            Top             =   6400
            Width           =   900
         End
         Begin VB.Label Label79 
            BackColor       =   &H00C0C0C0&
            Caption         =   "            Правый"
            Height          =   300
            Left            =   3600
            TabIndex        =   301
            Top             =   5900
            Width           =   2400
         End
         Begin VB.Label Label78 
            BackColor       =   &H00C0C0C0&
            Caption         =   "             Левый"
            Height          =   300
            Left            =   150
            TabIndex        =   300
            Top             =   5900
            Width           =   2400
         End
         Begin VB.Label Label77 
            BackColor       =   &H00C0C0C0&
            Caption         =   "                                    Сопряженка"
            Height          =   300
            Left            =   150
            TabIndex        =   299
            Top             =   5400
            Width           =   5850
         End
         Begin VB.Label Label76 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Мц="
            Height          =   300
            Left            =   150
            TabIndex        =   293
            Top             =   4700
            Width           =   500
         End
         Begin VB.Label Label75 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Д="
            Height          =   300
            Left            =   150
            TabIndex        =   292
            Top             =   4200
            Width           =   400
         End
         Begin VB.Label Label74 
            BackColor       =   &H00C0C0C0&
            Caption         =   "А="
            Height          =   300
            Left            =   150
            TabIndex        =   291
            Top             =   3700
            Width           =   400
         End
         Begin VB.Label Label73 
            BackColor       =   &H00C0C0C0&
            Caption         =   "№ КП="
            Height          =   300
            Index           =   0
            Left            =   150
            TabIndex        =   290
            Top             =   3200
            Width           =   800
         End
         Begin VB.Label Label72 
            BackColor       =   &H00C0C0C0&
            Caption         =   "                         А, Д"
            Height          =   300
            Left            =   150
            TabIndex        =   289
            Top             =   2700
            Width           =   3550
         End
         Begin VB.Label Label71 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   150
            TabIndex        =   284
            Top             =   1900
            Width           =   400
         End
         Begin VB.Label Label70 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   300
            Left            =   150
            TabIndex        =   283
            Top             =   1400
            Width           =   400
         End
         Begin VB.Label Label69 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Х="
            Height          =   300
            Left            =   150
            TabIndex        =   282
            Top             =   900
            Width           =   400
         End
         Begin VB.Label Label68 
            BackColor       =   &H00C0C0C0&
            Caption         =   "                          Х, У"
            Height          =   300
            Left            =   120
            TabIndex        =   281
            Top             =   405
            Width           =   3555
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Установки"
         Height          =   9600
         Left            =   -65880
         TabIndex        =   185
         Top             =   600
         Width           =   9200
         Begin VB.TextBox Text16 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   257
            Text            =   "0"
            Top             =   800
            Width           =   1500
         End
         Begin VB.TextBox Text17 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   256
            Text            =   "0"
            Top             =   1250
            Width           =   1500
         End
         Begin VB.TextBox Text18 
            BackColor       =   &H0080C0FF&
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
            Index           =   1
            Left            =   1800
            TabIndex        =   255
            Text            =   "0"
            Top             =   1700
            Width           =   1500
         End
         Begin VB.TextBox Text19 
            BackColor       =   &H0080C0FF&
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
            Index           =   1
            Left            =   1800
            TabIndex        =   254
            Text            =   "0"
            Top             =   2150
            Width           =   1500
         End
         Begin VB.TextBox Text20 
            BackColor       =   &H0080C0FF&
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
            Index           =   1
            Left            =   1800
            TabIndex        =   253
            Text            =   "0"
            Top             =   2600
            Width           =   1500
         End
         Begin VB.TextBox Text21 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   252
            Text            =   "0"
            Top             =   3050
            Width           =   1500
         End
         Begin VB.TextBox Text22 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   251
            Text            =   "0"
            Top             =   3500
            Width           =   1500
         End
         Begin VB.TextBox Text23 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   250
            Text            =   "0"
            Top             =   3950
            Width           =   1500
         End
         Begin VB.TextBox Text24 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   249
            Text            =   "0"
            Top             =   4400
            Width           =   1500
         End
         Begin VB.TextBox Text25 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   248
            Text            =   "0"
            Top             =   4850
            Width           =   1500
         End
         Begin VB.TextBox Text26 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   247
            Text            =   "0"
            Top             =   5300
            Width           =   1500
         End
         Begin VB.TextBox Text27 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   246
            Text            =   "0"
            Top             =   5750
            Width           =   1500
         End
         Begin VB.TextBox Text28 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   2
            Left            =   1800
            TabIndex        =   245
            Text            =   "0"
            Top             =   6200
            Width           =   1500
         End
         Begin VB.TextBox Text29 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   2
            Left            =   1800
            TabIndex        =   244
            Text            =   "0"
            Top             =   6650
            Width           =   1500
         End
         Begin VB.TextBox Text30 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   2
            Left            =   1800
            TabIndex        =   243
            Text            =   "0"
            Top             =   7100
            Width           =   1500
         End
         Begin VB.TextBox Text31 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   2
            Left            =   1800
            TabIndex        =   242
            Text            =   "0"
            Top             =   7550
            Width           =   1500
         End
         Begin VB.TextBox Text32 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   241
            Text            =   "0"
            Top             =   8000
            Width           =   1500
         End
         Begin VB.TextBox Text33 
            BackColor       =   &H0080C0FF&
            Height          =   400
            Index           =   1
            Left            =   1800
            TabIndex        =   240
            Text            =   "0"
            Top             =   8450
            Width           =   1500
         End
         Begin VB.TextBox Text34 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   239
            Text            =   "0"
            Top             =   800
            Width           =   1500
         End
         Begin VB.TextBox Text35 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   238
            Text            =   "0"
            Top             =   1250
            Width           =   1500
         End
         Begin VB.TextBox Text36 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   237
            Text            =   "0"
            Top             =   1700
            Width           =   1500
         End
         Begin VB.TextBox Text37 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   236
            Text            =   "0"
            Top             =   2150
            Width           =   1500
         End
         Begin VB.TextBox Text38 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   235
            Text            =   "0"
            Top             =   2600
            Width           =   1500
         End
         Begin VB.TextBox Text39 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   234
            Text            =   "0"
            Top             =   3050
            Width           =   1500
         End
         Begin VB.TextBox Text40 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   233
            Text            =   "0"
            Top             =   3500
            Width           =   1500
         End
         Begin VB.TextBox Text41 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   232
            Text            =   "0"
            Top             =   3950
            Width           =   1500
         End
         Begin VB.TextBox Text42 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   231
            Text            =   "0"
            Top             =   4400
            Width           =   1500
         End
         Begin VB.TextBox Text43 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   230
            Text            =   "0"
            Top             =   4850
            Width           =   1500
         End
         Begin VB.TextBox Text44 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   229
            Text            =   "0"
            Top             =   800
            Width           =   1500
         End
         Begin VB.TextBox Text45 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   228
            Text            =   "0"
            Top             =   1250
            Width           =   1500
         End
         Begin VB.TextBox Text46 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   227
            Text            =   "0"
            Top             =   1700
            Width           =   1500
         End
         Begin VB.TextBox Text47 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   226
            Text            =   "0"
            Top             =   2150
            Width           =   1500
         End
         Begin VB.TextBox Text48 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   225
            Text            =   "0"
            Top             =   2600
            Width           =   1500
         End
         Begin VB.TextBox Text49 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   224
            Text            =   "0"
            Top             =   3050
            Width           =   1500
         End
         Begin VB.TextBox Text50 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   223
            Text            =   "0"
            Top             =   3500
            Width           =   1500
         End
         Begin VB.TextBox Text51 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   222
            Text            =   "0"
            Top             =   3950
            Width           =   1500
         End
         Begin VB.TextBox Text52 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   221
            Text            =   "0"
            Top             =   4400
            Width           =   1500
         End
         Begin VB.TextBox Text53 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   220
            Text            =   "0"
            Top             =   4850
            Width           =   1500
         End
         Begin VB.TextBox Text54 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   219
            Text            =   "0"
            Top             =   7100
            Width           =   1500
         End
         Begin VB.TextBox Text55 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   218
            Text            =   "0"
            Top             =   800
            Width           =   1500
         End
         Begin VB.TextBox Text56 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   217
            Text            =   "0"
            Top             =   1250
            Width           =   1500
         End
         Begin VB.TextBox Text57 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   216
            Text            =   "0"
            Top             =   1700
            Width           =   1500
         End
         Begin VB.TextBox Text58 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   215
            Text            =   "0"
            Top             =   2150
            Width           =   1500
         End
         Begin VB.TextBox Text59 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   214
            Text            =   "0"
            Top             =   2600
            Width           =   1500
         End
         Begin VB.TextBox Text60 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   213
            Text            =   "0"
            Top             =   3050
            Width           =   1500
         End
         Begin VB.TextBox Text61 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   212
            Text            =   "0"
            Top             =   3500
            Width           =   1500
         End
         Begin VB.TextBox Text62 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   211
            Text            =   "0"
            Top             =   3950
            Width           =   1500
         End
         Begin VB.TextBox Text63 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   210
            Text            =   "0"
            Top             =   4400
            Width           =   1500
         End
         Begin VB.TextBox Text64 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   209
            Text            =   "0"
            Top             =   4850
            Width           =   1500
         End
         Begin VB.TextBox Text65 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   208
            Text            =   "0"
            Top             =   5300
            Width           =   1500
         End
         Begin VB.TextBox Text66 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   207
            Text            =   "0"
            Top             =   5750
            Width           =   1500
         End
         Begin VB.TextBox Text67 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   206
            Text            =   "0"
            Top             =   6200
            Width           =   1500
         End
         Begin VB.TextBox Text68 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   205
            Text            =   "0"
            Top             =   6650
            Width           =   1500
         End
         Begin VB.TextBox Text69 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   204
            Text            =   "0"
            Top             =   7100
            Width           =   1500
         End
         Begin VB.TextBox Text70 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   203
            Text            =   "0"
            Top             =   8450
            Width           =   1500
         End
         Begin VB.TextBox Text71 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   202
            Text            =   "0"
            Top             =   5300
            Width           =   1500
         End
         Begin VB.TextBox Text72 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   201
            Text            =   "0"
            Top             =   5750
            Width           =   1500
         End
         Begin VB.TextBox Text73 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   200
            Text            =   "0"
            Top             =   6200
            Width           =   1500
         End
         Begin VB.TextBox Text74 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   199
            Text            =   "0"
            Top             =   6650
            Width           =   1500
         End
         Begin VB.TextBox Text75 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   198
            Text            =   "0"
            Top             =   5300
            Width           =   1500
         End
         Begin VB.TextBox Text76 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   197
            Text            =   "0"
            Top             =   5750
            Width           =   1500
         End
         Begin VB.TextBox Text77 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   196
            Text            =   "0"
            Top             =   6200
            Width           =   1500
         End
         Begin VB.TextBox Text78 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   195
            Text            =   "0"
            Top             =   6650
            Width           =   1500
         End
         Begin VB.TextBox Text79 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   194
            Text            =   "0"
            Top             =   7100
            Width           =   1500
         End
         Begin VB.TextBox Text80 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Index           =   1
            Left            =   6900
            TabIndex        =   193
            Text            =   "0"
            Top             =   8450
            Width           =   1500
         End
         Begin VB.TextBox Text81 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   192
            Text            =   "0"
            Top             =   7550
            Width           =   1500
         End
         Begin VB.TextBox Text82 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00000080&
            Height          =   400
            Index           =   1
            Left            =   3500
            TabIndex        =   191
            Text            =   "0"
            Top             =   8000
            Width           =   1500
         End
         Begin VB.TextBox Text83 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   190
            Text            =   "0"
            Top             =   7550
            Width           =   1500
         End
         Begin VB.TextBox Text84 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Index           =   1
            Left            =   5200
            TabIndex        =   189
            Text            =   "0"
            Top             =   8000
            Width           =   1500
         End
         Begin VB.TextBox Text96 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00008000&
            Height          =   400
            Left            =   5200
            TabIndex        =   188
            Text            =   "0"
            Top             =   8450
            Width           =   1500
         End
         Begin VB.TextBox Text95 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Left            =   6900
            TabIndex        =   187
            Text            =   "0"
            Top             =   7550
            Width           =   1500
         End
         Begin VB.TextBox Text94 
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H00C00000&
            Height          =   400
            Left            =   6900
            TabIndex        =   186
            Text            =   "0"
            Top             =   8000
            Width           =   1500
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0C0C0&
            Caption         =   "       1 Бат"
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
            Index           =   1
            Left            =   1800
            TabIndex        =   279
            Top             =   400
            Width           =   1500
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            Caption         =   "       2 Бат"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Index           =   1
            Left            =   3500
            TabIndex        =   278
            Top             =   400
            Width           =   1500
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C0C0&
            Caption         =   "       3 Бат"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   300
            Index           =   1
            Left            =   5200
            TabIndex        =   277
            Top             =   400
            Width           =   1500
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            Caption         =   "       4 Бат"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   1
            Left            =   6900
            TabIndex        =   276
            Top             =   400
            Width           =   1500
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Снаряд"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   275
            Top             =   800
            Width           =   1500
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Взрыватель"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   274
            Top             =   1250
            Width           =   1500
         End
         Begin VB.Label Label29 
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
            Index           =   1
            Left            =   100
            TabIndex        =   273
            Top             =   1700
            Width           =   1500
         End
         Begin VB.Label Label30 
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
            Index           =   2
            Left            =   100
            TabIndex        =   272
            Top             =   2150
            Width           =   1500
         End
         Begin VB.Label Label31 
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
            Index           =   2
            Left            =   100
            TabIndex        =   271
            Top             =   2600
            Width           =   1700
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Веер"
            Height          =   300
            Index           =   2
            Left            =   100
            TabIndex        =   270
            Top             =   3050
            Width           =   1500
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Скачок"
            Height          =   300
            Index           =   2
            Left            =   100
            TabIndex        =   269
            Top             =   3500
            Width           =   1500
         End
         Begin VB.Label Label34 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Дал. Топо."
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   268
            Top             =   3950
            Width           =   1500
         End
         Begin VB.Label Label35 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Довор. Топо"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   267
            Top             =   4400
            Width           =   1500
         End
         Begin VB.Label Label36 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Мц"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   266
            Top             =   4850
            Width           =   1500
         End
         Begin VB.Label Label37 
            BackColor       =   &H00C0C0C0&
            Caption         =   "dXтыс"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   265
            Top             =   5300
            Width           =   1500
         End
         Begin VB.Label Label38 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Полетное"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   264
            Top             =   5750
            Width           =   1500
         End
         Begin VB.Label Label39 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Вд"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   263
            Top             =   6200
            Width           =   1500
         End
         Begin VB.Label Label40 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Выс. траект."
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   262
            Top             =   6650
            Width           =   1500
         End
         Begin VB.Label Label41 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ОН"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   261
            Top             =   7100
            Width           =   1500
         End
         Begin VB.Label Label42 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Хц"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   260
            Top             =   7550
            Width           =   1500
         End
         Begin VB.Label Label43 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Уц"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   259
            Top             =   8000
            Width           =   1500
         End
         Begin VB.Label Label44 
            BackColor       =   &H00C0C0C0&
            Caption         =   "hц"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   258
            Top             =   8450
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Боевой порядок"
         Height          =   9500
         Left            =   50
         TabIndex        =   120
         Top             =   500
         Width           =   8600
         Begin VB.TextBox pOH4 
            Height          =   400
            Index           =   0
            Left            =   7100
            TabIndex        =   151
            Text            =   "0"
            Top             =   3100
            Width           =   1000
         End
         Begin VB.TextBox pOH3 
            Height          =   400
            Index           =   0
            Left            =   7100
            TabIndex        =   150
            Text            =   "0"
            Top             =   2400
            Width           =   1000
         End
         Begin VB.TextBox pOH2 
            Height          =   400
            Index           =   0
            Left            =   7100
            TabIndex        =   149
            Text            =   "0"
            Top             =   1700
            Width           =   1000
         End
         Begin VB.TextBox pOH1 
            Height          =   400
            Index           =   0
            Left            =   7100
            TabIndex        =   148
            Text            =   "0"
            Top             =   1000
            Width           =   1000
         End
         Begin VB.TextBox pX1 
            Height          =   400
            Left            =   1200
            TabIndex        =   147
            Text            =   "0"
            Top             =   1000
            Width           =   1500
         End
         Begin VB.TextBox pY1 
            Height          =   400
            Left            =   3200
            TabIndex        =   146
            Text            =   "0"
            Top             =   1000
            Width           =   1500
         End
         Begin VB.TextBox ph1 
            Height          =   400
            Left            =   5200
            TabIndex        =   145
            Text            =   "0"
            Top             =   1000
            Width           =   1000
         End
         Begin VB.TextBox pX2 
            Height          =   400
            Left            =   1200
            TabIndex        =   144
            Text            =   "0"
            Top             =   1700
            Width           =   1500
         End
         Begin VB.TextBox pY2 
            Height          =   400
            Left            =   3200
            TabIndex        =   143
            Text            =   "0"
            Top             =   1700
            Width           =   1500
         End
         Begin VB.TextBox ph2 
            Height          =   400
            Left            =   5200
            TabIndex        =   142
            Text            =   "0"
            Top             =   1700
            Width           =   1000
         End
         Begin VB.TextBox pX3 
            Height          =   400
            Left            =   1200
            TabIndex        =   141
            Text            =   "0"
            Top             =   2400
            Width           =   1500
         End
         Begin VB.TextBox pY3 
            Height          =   400
            Left            =   3200
            TabIndex        =   140
            Text            =   "0"
            Top             =   2400
            Width           =   1500
         End
         Begin VB.TextBox ph3 
            Height          =   400
            Left            =   5200
            TabIndex        =   139
            Text            =   "0"
            Top             =   2400
            Width           =   1000
         End
         Begin VB.TextBox pX4 
            Height          =   400
            Left            =   1200
            TabIndex        =   138
            Text            =   "0"
            Top             =   3100
            Width           =   1500
         End
         Begin VB.TextBox pY4 
            Height          =   400
            Left            =   3200
            TabIndex        =   137
            Text            =   "0"
            Top             =   3100
            Width           =   1500
         End
         Begin VB.TextBox ph4 
            Height          =   400
            Left            =   5200
            TabIndex        =   136
            Text            =   "0"
            Top             =   3100
            Width           =   1000
         End
         Begin VB.TextBox pXkp1 
            Height          =   400
            Index           =   0
            Left            =   1200
            TabIndex        =   135
            Text            =   "0"
            Top             =   4300
            Width           =   1500
         End
         Begin VB.TextBox pYkp1 
            Height          =   400
            Index           =   1
            Left            =   3200
            TabIndex        =   134
            Text            =   "0"
            Top             =   4300
            Width           =   1500
         End
         Begin VB.TextBox phkp1 
            Height          =   400
            Left            =   5200
            TabIndex        =   133
            Text            =   "0"
            Top             =   4300
            Width           =   1000
         End
         Begin VB.TextBox pXkp2 
            Height          =   400
            Index           =   0
            Left            =   1200
            TabIndex        =   132
            Text            =   "0"
            Top             =   5000
            Width           =   1500
         End
         Begin VB.TextBox pTkp2 
            Height          =   400
            Index           =   0
            Left            =   3200
            TabIndex        =   131
            Text            =   "0"
            Top             =   5000
            Width           =   1500
         End
         Begin VB.TextBox phkp2 
            Height          =   400
            Index           =   0
            Left            =   5200
            TabIndex        =   130
            Text            =   "0"
            Top             =   5000
            Width           =   1000
         End
         Begin VB.TextBox pXkp3 
            Height          =   400
            Index           =   0
            Left            =   1200
            TabIndex        =   129
            Text            =   "0"
            Top             =   5700
            Width           =   1500
         End
         Begin VB.TextBox pYkp3 
            Height          =   400
            Index           =   0
            Left            =   3200
            TabIndex        =   128
            Text            =   "0"
            Top             =   5700
            Width           =   1500
         End
         Begin VB.TextBox phkp3 
            Height          =   400
            Index           =   0
            Left            =   5200
            TabIndex        =   127
            Text            =   "0"
            Top             =   5700
            Width           =   1000
         End
         Begin VB.TextBox pXkp4 
            Height          =   400
            Index           =   0
            Left            =   1200
            TabIndex        =   126
            Text            =   "0"
            Top             =   6400
            Width           =   1500
         End
         Begin VB.TextBox pYkp4 
            Height          =   400
            Index           =   0
            Left            =   3200
            TabIndex        =   125
            Text            =   "0"
            Top             =   6400
            Width           =   1500
         End
         Begin VB.TextBox phkp4 
            Height          =   400
            Index           =   0
            Left            =   5200
            TabIndex        =   124
            Text            =   "0"
            Top             =   6400
            Width           =   1000
         End
         Begin VB.TextBox pXkp5 
            Height          =   400
            Index           =   0
            Left            =   1200
            TabIndex        =   123
            Text            =   "0"
            Top             =   7100
            Width           =   1500
         End
         Begin VB.TextBox pYkp5 
            Height          =   400
            Index           =   0
            Left            =   3200
            TabIndex        =   122
            Text            =   "0"
            Top             =   7100
            Width           =   1500
         End
         Begin VB.TextBox phkp5 
            Height          =   400
            Index           =   0
            Left            =   5200
            TabIndex        =   121
            Text            =   "0"
            Top             =   7100
            Width           =   1000
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ОН="
            Height          =   400
            Index           =   0
            Left            =   6500
            TabIndex        =   184
            Top             =   3100
            Width           =   600
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ОН="
            Height          =   400
            Index           =   0
            Left            =   6500
            TabIndex        =   183
            Top             =   2400
            Width           =   600
         End
         Begin VB.Label Label31 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ОН="
            Height          =   400
            Index           =   0
            Left            =   6500
            TabIndex        =   182
            Top             =   1700
            Width           =   600
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ОН="
            Height          =   400
            Index           =   0
            Left            =   6500
            TabIndex        =   181
            Top             =   1000
            Width           =   600
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "                                                       ОГНЕВЫЕ"
            Height          =   300
            Left            =   300
            TabIndex        =   180
            Top             =   480
            Width           =   7800
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1 Бат Х="
            Height          =   300
            Left            =   200
            TabIndex        =   179
            Top             =   1000
            Width           =   1000
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   405
            Left            =   2800
            TabIndex        =   178
            Top             =   1005
            Width           =   300
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   4800
            TabIndex        =   177
            Top             =   1005
            Width           =   300
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2 Бат Х="
            Height          =   300
            Left            =   200
            TabIndex        =   176
            Top             =   1700
            Width           =   1000
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   400
            Left            =   2800
            TabIndex        =   175
            Top             =   1700
            Width           =   300
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   4800
            TabIndex        =   174
            Top             =   1700
            Width           =   300
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3 Бат Х="
            Height          =   300
            Left            =   200
            TabIndex        =   173
            Top             =   2400
            Width           =   1000
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "                                               КНП"
            Height          =   300
            Left            =   300
            TabIndex        =   172
            Top             =   3800
            Width           =   6200
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            Caption         =   "КНП 1 Х="
            Height          =   300
            Left            =   100
            TabIndex        =   171
            Top             =   4300
            Width           =   1100
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   400
            Left            =   2800
            TabIndex        =   170
            Top             =   4300
            Width           =   300
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   400
            Left            =   2800
            TabIndex        =   169
            Top             =   2400
            Width           =   300
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   4800
            TabIndex        =   168
            Top             =   2400
            Width           =   300
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0C0C0&
            Caption         =   "4 Бат Х="
            Height          =   300
            Left            =   200
            TabIndex        =   167
            Top             =   3100
            Width           =   1000
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   400
            Left            =   2800
            TabIndex        =   166
            Top             =   3100
            Width           =   300
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   4800
            TabIndex        =   165
            Top             =   3100
            Width           =   300
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   4800
            TabIndex        =   164
            Top             =   4300
            Width           =   300
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "КНП 2 Х="
            Height          =   300
            Left            =   100
            TabIndex        =   163
            Top             =   5000
            Width           =   1100
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   400
            Left            =   2800
            TabIndex        =   162
            Top             =   5000
            Width           =   300
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   4800
            TabIndex        =   161
            Top             =   5000
            Width           =   300
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            Caption         =   "КНП 3 Х="
            Height          =   300
            Left            =   100
            TabIndex        =   160
            Top             =   5700
            Width           =   1100
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   400
            Left            =   2800
            TabIndex        =   159
            Top             =   5700
            Width           =   300
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Index           =   0
            Left            =   4800
            TabIndex        =   158
            Top             =   5700
            Width           =   300
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            Caption         =   "КНП 4 Х="
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   157
            Top             =   6405
            Width           =   1095
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   400
            Index           =   0
            Left            =   2800
            TabIndex        =   156
            Top             =   6400
            Width           =   300
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Index           =   0
            Left            =   4800
            TabIndex        =   155
            Top             =   6400
            Width           =   300
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0C0C0&
            Caption         =   "КНП 5 Х="
            Height          =   300
            Index           =   0
            Left            =   100
            TabIndex        =   154
            Top             =   7100
            Width           =   1100
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0C0C0&
            Caption         =   "У="
            Height          =   400
            Index           =   0
            Left            =   2800
            TabIndex        =   153
            Top             =   7100
            Width           =   300
         End
         Begin VB.Label Label29 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Index           =   0
            Left            =   4800
            TabIndex        =   152
            Top             =   7100
            Width           =   300
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Балистика"
         Height          =   6600
         Left            =   8760
         TabIndex        =   79
         Top             =   500
         Width           =   5300
         Begin VB.TextBox pTz1 
            Height          =   400
            Index           =   1
            Left            =   1100
            TabIndex        =   107
            Text            =   "0"
            Top             =   1300
            Width           =   600
         End
         Begin VB.TextBox pTz2 
            Height          =   400
            Index           =   1
            Left            =   2100
            TabIndex        =   106
            Text            =   "0"
            Top             =   1300
            Width           =   600
         End
         Begin VB.TextBox pTz3 
            Height          =   400
            Index           =   1
            Left            =   3100
            TabIndex        =   105
            Text            =   "0"
            Top             =   1300
            Width           =   600
         End
         Begin VB.TextBox pTz4 
            Height          =   400
            Index           =   1
            Left            =   4100
            TabIndex        =   104
            Text            =   "0"
            Top             =   1300
            Width           =   600
         End
         Begin VB.TextBox pV0p1 
            Height          =   400
            Index           =   0
            Left            =   1100
            TabIndex        =   103
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox pV0p2 
            Height          =   400
            Index           =   0
            Left            =   2100
            TabIndex        =   102
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox pV0p3 
            Height          =   400
            Index           =   0
            Left            =   3100
            TabIndex        =   101
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox pV0p4 
            Height          =   400
            Index           =   0
            Left            =   4100
            TabIndex        =   100
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox pV0y1 
            Height          =   400
            Index           =   0
            Left            =   1100
            TabIndex        =   99
            Text            =   "0"
            Top             =   2900
            Width           =   600
         End
         Begin VB.TextBox pV0y2 
            Height          =   400
            Index           =   0
            Left            =   2100
            TabIndex        =   98
            Text            =   "0"
            Top             =   2900
            Width           =   600
         End
         Begin VB.TextBox pV0y3 
            Height          =   400
            Index           =   0
            Left            =   3100
            TabIndex        =   97
            Text            =   "0"
            Top             =   2900
            Width           =   600
         End
         Begin VB.TextBox pV0y4 
            Height          =   400
            Index           =   0
            Left            =   4100
            TabIndex        =   96
            Text            =   "0"
            Top             =   2900
            Width           =   600
         End
         Begin VB.TextBox pV011 
            Height          =   400
            Index           =   0
            Left            =   1100
            TabIndex        =   95
            Text            =   "0"
            Top             =   3500
            Width           =   600
         End
         Begin VB.TextBox pV012 
            Height          =   400
            Index           =   0
            Left            =   2100
            TabIndex        =   94
            Text            =   "0"
            Top             =   3500
            Width           =   600
         End
         Begin VB.TextBox pV013 
            Height          =   400
            Index           =   0
            Left            =   3100
            TabIndex        =   93
            Text            =   "0"
            Top             =   3500
            Width           =   600
         End
         Begin VB.TextBox pV014 
            Height          =   400
            Index           =   0
            Left            =   4100
            TabIndex        =   92
            Text            =   "0"
            Top             =   3500
            Width           =   600
         End
         Begin VB.TextBox pV021 
            Height          =   400
            Index           =   0
            Left            =   1100
            TabIndex        =   91
            Text            =   "0"
            Top             =   4100
            Width           =   600
         End
         Begin VB.TextBox pV022 
            Height          =   400
            Index           =   0
            Left            =   2100
            TabIndex        =   90
            Text            =   "0"
            Top             =   4100
            Width           =   600
         End
         Begin VB.TextBox pV023 
            Height          =   400
            Index           =   0
            Left            =   3100
            TabIndex        =   89
            Text            =   "0"
            Top             =   4100
            Width           =   600
         End
         Begin VB.TextBox pV024 
            Height          =   400
            Index           =   0
            Left            =   4100
            TabIndex        =   88
            Text            =   "0"
            Top             =   4100
            Width           =   600
         End
         Begin VB.TextBox pV031 
            Height          =   400
            Index           =   0
            Left            =   1100
            TabIndex        =   87
            Text            =   "0"
            Top             =   4700
            Width           =   600
         End
         Begin VB.TextBox pV032 
            Height          =   400
            Index           =   0
            Left            =   2100
            TabIndex        =   86
            Text            =   "0"
            Top             =   4700
            Width           =   600
         End
         Begin VB.TextBox pV033 
            Height          =   400
            Index           =   0
            Left            =   3100
            TabIndex        =   85
            Text            =   "0"
            Top             =   4700
            Width           =   600
         End
         Begin VB.TextBox pV034 
            Height          =   400
            Index           =   0
            Left            =   4100
            TabIndex        =   84
            Text            =   "0"
            Top             =   4700
            Width           =   600
         End
         Begin VB.TextBox pV041 
            Height          =   400
            Index           =   0
            Left            =   1100
            TabIndex        =   83
            Text            =   "0"
            Top             =   5300
            Width           =   600
         End
         Begin VB.TextBox pV042 
            Height          =   400
            Index           =   0
            Left            =   2100
            TabIndex        =   82
            Text            =   "0"
            Top             =   5300
            Width           =   600
         End
         Begin VB.TextBox pV043 
            Height          =   400
            Index           =   0
            Left            =   3100
            TabIndex        =   81
            Text            =   "0"
            Top             =   5300
            Width           =   600
         End
         Begin VB.TextBox pV044 
            Height          =   400
            Index           =   0
            Left            =   4100
            TabIndex        =   80
            Text            =   "0"
            Top             =   5300
            Width           =   600
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C0C0C0&
            Caption         =   "            Температура заряда"
            Height          =   300
            Index           =   1
            Left            =   1100
            TabIndex        =   119
            Top             =   900
            Width           =   3600
         End
         Begin VB.Label Label31 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1 Бат"
            Height          =   300
            Index           =   1
            Left            =   1100
            TabIndex        =   118
            Top             =   500
            Width           =   600
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2 Бат"
            Height          =   300
            Index           =   1
            Left            =   2100
            TabIndex        =   117
            Top             =   500
            Width           =   600
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3 Бат"
            Height          =   300
            Index           =   1
            Left            =   3100
            TabIndex        =   116
            Top             =   500
            Width           =   600
         End
         Begin VB.Label Label34 
            BackColor       =   &H00C0C0C0&
            Caption         =   "4 Бат"
            Height          =   300
            Index           =   0
            Left            =   4100
            TabIndex        =   115
            Top             =   500
            Width           =   600
         End
         Begin VB.Label Label35 
            BackColor       =   &H00C0C0C0&
            Caption         =   "       Потеря нач. скорости Vо"
            Height          =   300
            Index           =   0
            Left            =   1100
            TabIndex        =   114
            Top             =   1900
            Width           =   3600
         End
         Begin VB.Label Label36 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Полн"
            Height          =   300
            Index           =   0
            Left            =   400
            TabIndex        =   113
            Top             =   2300
            Width           =   600
         End
         Begin VB.Label Label37 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Умен"
            Height          =   300
            Index           =   0
            Left            =   400
            TabIndex        =   112
            Top             =   2900
            Width           =   600
         End
         Begin VB.Label Label38 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Перв"
            Height          =   300
            Index           =   0
            Left            =   400
            TabIndex        =   111
            Top             =   3500
            Width           =   600
         End
         Begin VB.Label Label39 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Втор"
            Height          =   300
            Index           =   0
            Left            =   400
            TabIndex        =   110
            Top             =   4100
            Width           =   600
         End
         Begin VB.Label Label40 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Трет"
            Height          =   300
            Index           =   0
            Left            =   400
            TabIndex        =   109
            Top             =   4700
            Width           =   600
         End
         Begin VB.Label Label41 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Четв"
            Height          =   300
            Index           =   0
            Left            =   400
            TabIndex        =   108
            Top             =   5300
            Width           =   600
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Метеоданные"
         Height          =   6600
         Left            =   14160
         TabIndex        =   21
         Top             =   500
         Width           =   5700
         Begin VB.TextBox phMet 
            Height          =   400
            Index           =   0
            Left            =   1300
            TabIndex        =   59
            Text            =   "0"
            Top             =   500
            Width           =   1000
         End
         Begin VB.TextBox pH 
            Height          =   400
            Index           =   0
            Left            =   1300
            TabIndex        =   58
            Text            =   "0"
            Top             =   1000
            Width           =   1000
         End
         Begin VB.TextBox pTz 
            Height          =   400
            Index           =   0
            Left            =   1300
            TabIndex        =   57
            Text            =   "0"
            Top             =   1500
            Width           =   1000
         End
         Begin VB.TextBox pAw 
            Height          =   400
            Index           =   0
            Left            =   1300
            TabIndex        =   56
            Text            =   "0"
            Top             =   2000
            Width           =   1000
         End
         Begin VB.TextBox pW 
            Height          =   400
            Index           =   0
            Left            =   1300
            TabIndex        =   55
            Text            =   "0"
            Top             =   2500
            Width           =   1000
         End
         Begin VB.TextBox pdt02 
            Height          =   400
            Index           =   0
            Left            =   3000
            TabIndex        =   54
            Text            =   "0"
            Top             =   800
            Width           =   600
         End
         Begin VB.TextBox pAw02 
            Height          =   400
            Index           =   0
            Left            =   3800
            TabIndex        =   53
            Text            =   "0"
            Top             =   800
            Width           =   600
         End
         Begin VB.TextBox pW02 
            Height          =   400
            Index           =   0
            Left            =   4600
            TabIndex        =   52
            Text            =   "0"
            Top             =   800
            Width           =   600
         End
         Begin VB.TextBox pdt04 
            Height          =   400
            Index           =   0
            Left            =   3000
            TabIndex        =   51
            Text            =   "0"
            Top             =   1300
            Width           =   600
         End
         Begin VB.TextBox pAw04 
            Height          =   400
            Index           =   0
            Left            =   3800
            TabIndex        =   50
            Text            =   "0"
            Top             =   1300
            Width           =   600
         End
         Begin VB.TextBox pW04 
            Height          =   400
            Index           =   0
            Left            =   4600
            TabIndex        =   49
            Text            =   "0"
            Top             =   1300
            Width           =   600
         End
         Begin VB.TextBox pdt08 
            Height          =   400
            Index           =   0
            Left            =   3000
            TabIndex        =   48
            Text            =   "0"
            Top             =   1800
            Width           =   600
         End
         Begin VB.TextBox pdt12 
            Height          =   400
            Index           =   0
            Left            =   3000
            TabIndex        =   47
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox pdt16 
            Height          =   400
            Index           =   0
            Left            =   3000
            TabIndex        =   46
            Text            =   "0"
            Top             =   2800
            Width           =   600
         End
         Begin VB.TextBox pdt20 
            Height          =   400
            Index           =   0
            Left            =   3000
            TabIndex        =   45
            Text            =   "0"
            Top             =   3300
            Width           =   600
         End
         Begin VB.TextBox pdt24 
            Height          =   400
            Index           =   0
            Left            =   3000
            TabIndex        =   44
            Text            =   "0"
            Top             =   3800
            Width           =   600
         End
         Begin VB.TextBox pdt30 
            Height          =   400
            Index           =   0
            Left            =   3000
            TabIndex        =   43
            Text            =   "0"
            Top             =   4300
            Width           =   600
         End
         Begin VB.TextBox pAw08 
            Height          =   400
            Index           =   0
            Left            =   3800
            TabIndex        =   42
            Text            =   "0"
            Top             =   1800
            Width           =   600
         End
         Begin VB.TextBox pAw12 
            Height          =   400
            Index           =   0
            Left            =   3800
            TabIndex        =   41
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox pAw16 
            Height          =   400
            Index           =   0
            Left            =   3800
            TabIndex        =   40
            Text            =   "0"
            Top             =   2800
            Width           =   600
         End
         Begin VB.TextBox pAw20 
            Height          =   400
            Index           =   0
            Left            =   3800
            TabIndex        =   39
            Text            =   "0"
            Top             =   3300
            Width           =   600
         End
         Begin VB.TextBox pAw24 
            Height          =   400
            Index           =   0
            Left            =   3800
            TabIndex        =   38
            Text            =   "0"
            Top             =   3800
            Width           =   600
         End
         Begin VB.TextBox pAw30 
            Height          =   400
            Index           =   0
            Left            =   3800
            TabIndex        =   37
            Text            =   "0"
            Top             =   4300
            Width           =   600
         End
         Begin VB.TextBox pW08 
            Height          =   400
            Index           =   0
            Left            =   4600
            TabIndex        =   36
            Text            =   "0"
            Top             =   1800
            Width           =   600
         End
         Begin VB.TextBox pW12 
            Height          =   400
            Index           =   0
            Left            =   4600
            TabIndex        =   35
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox pW16 
            Height          =   400
            Index           =   0
            Left            =   4600
            TabIndex        =   34
            Text            =   "0"
            Top             =   2800
            Width           =   600
         End
         Begin VB.TextBox pW20 
            Height          =   400
            Index           =   0
            Left            =   4600
            TabIndex        =   33
            Text            =   "0"
            Top             =   3300
            Width           =   600
         End
         Begin VB.TextBox pW24 
            Height          =   400
            Index           =   0
            Left            =   4600
            TabIndex        =   32
            Text            =   "0"
            Top             =   3800
            Width           =   600
         End
         Begin VB.TextBox pW30 
            Height          =   400
            Index           =   0
            Left            =   4600
            TabIndex        =   31
            Text            =   "0"
            Top             =   4300
            Width           =   600
         End
         Begin VB.TextBox pdt40 
            Height          =   400
            Left            =   3000
            TabIndex        =   30
            Text            =   "0"
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox pdt50 
            Height          =   400
            Left            =   3000
            TabIndex        =   29
            Text            =   "0"
            Top             =   5300
            Width           =   600
         End
         Begin VB.TextBox pdt60 
            Height          =   400
            Left            =   3000
            TabIndex        =   28
            Text            =   "0"
            Top             =   5800
            Width           =   600
         End
         Begin VB.TextBox pAw40 
            Height          =   400
            Left            =   3800
            TabIndex        =   27
            Text            =   "0"
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox pAw50 
            Height          =   400
            Left            =   3800
            TabIndex        =   26
            Text            =   "0"
            Top             =   5300
            Width           =   600
         End
         Begin VB.TextBox pAw60 
            Height          =   400
            Left            =   3800
            TabIndex        =   25
            Text            =   "0"
            Top             =   5800
            Width           =   600
         End
         Begin VB.TextBox pW40 
            Height          =   400
            Left            =   4600
            TabIndex        =   24
            Text            =   "0"
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox pW50 
            Height          =   400
            Left            =   4600
            TabIndex        =   23
            Text            =   "0"
            Top             =   5300
            Width           =   600
         End
         Begin VB.TextBox pW60 
            Height          =   400
            Left            =   4600
            TabIndex        =   22
            Text            =   "0"
            Top             =   5800
            Width           =   600
         End
         Begin VB.Label Label42 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h метео="
            Height          =   300
            Index           =   0
            Left            =   200
            TabIndex        =   78
            Top             =   500
            Width           =   1100
         End
         Begin VB.Label Label43 
            BackColor       =   &H00C0C0C0&
            Caption         =   "H="
            Height          =   300
            Index           =   0
            Left            =   850
            TabIndex        =   77
            Top             =   1000
            Width           =   300
         End
         Begin VB.Label Label44 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tз="
            Height          =   300
            Index           =   0
            Left            =   750
            TabIndex        =   76
            Top             =   1500
            Width           =   400
         End
         Begin VB.Label Label45 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Aw="
            Height          =   300
            Left            =   700
            TabIndex        =   75
            Top             =   2000
            Width           =   400
         End
         Begin VB.Label Label46 
            BackColor       =   &H00C0C0C0&
            Caption         =   "W="
            Height          =   300
            Left            =   700
            TabIndex        =   74
            Top             =   2500
            Width           =   400
         End
         Begin VB.Label Label47 
            BackColor       =   &H00C0C0C0&
            Caption         =   "dT"
            Height          =   300
            Left            =   3000
            TabIndex        =   73
            Top             =   360
            Width           =   400
         End
         Begin VB.Label Label48 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Aw"
            Height          =   300
            Left            =   3800
            TabIndex        =   72
            Top             =   360
            Width           =   400
         End
         Begin VB.Label Label49 
            BackColor       =   &H00C0C0C0&
            Caption         =   "W"
            Height          =   300
            Left            =   4600
            TabIndex        =   71
            Top             =   360
            Width           =   400
         End
         Begin VB.Label Label50 
            BackColor       =   &H00C0C0C0&
            Caption         =   "02"
            Height          =   300
            Left            =   2600
            TabIndex        =   70
            Top             =   800
            Width           =   300
         End
         Begin VB.Label Label51 
            BackColor       =   &H00C0C0C0&
            Caption         =   "04"
            Height          =   300
            Left            =   2600
            TabIndex        =   69
            Top             =   1300
            Width           =   300
         End
         Begin VB.Label Label52 
            BackColor       =   &H00C0C0C0&
            Caption         =   "08"
            Height          =   300
            Left            =   2600
            TabIndex        =   68
            Top             =   1800
            Width           =   300
         End
         Begin VB.Label Label53 
            BackColor       =   &H00C0C0C0&
            Caption         =   "12"
            Height          =   300
            Left            =   2600
            TabIndex        =   67
            Top             =   2300
            Width           =   300
         End
         Begin VB.Label Label54 
            BackColor       =   &H00C0C0C0&
            Caption         =   "16"
            Height          =   300
            Left            =   2600
            TabIndex        =   66
            Top             =   2800
            Width           =   300
         End
         Begin VB.Label Label55 
            BackColor       =   &H00C0C0C0&
            Caption         =   "20"
            Height          =   300
            Left            =   2600
            TabIndex        =   65
            Top             =   3300
            Width           =   300
         End
         Begin VB.Label Label56 
            BackColor       =   &H00C0C0C0&
            Caption         =   "24"
            Height          =   300
            Left            =   2600
            TabIndex        =   64
            Top             =   3800
            Width           =   300
         End
         Begin VB.Label Label57 
            BackColor       =   &H00C0C0C0&
            Caption         =   "30"
            Height          =   300
            Left            =   2600
            TabIndex        =   63
            Top             =   4300
            Width           =   300
         End
         Begin VB.Label Label58 
            BackColor       =   &H00C0C0C0&
            Caption         =   "40"
            Height          =   300
            Left            =   2600
            TabIndex        =   62
            Top             =   4800
            Width           =   300
         End
         Begin VB.Label Label59 
            BackColor       =   &H00C0C0C0&
            Caption         =   "50"
            Height          =   300
            Left            =   2600
            TabIndex        =   61
            Top             =   5300
            Width           =   300
         End
         Begin VB.Label Label60 
            BackColor       =   &H00C0C0C0&
            Caption         =   "60"
            Height          =   300
            Left            =   2600
            TabIndex        =   60
            Top             =   5800
            Width           =   300
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Выбор Снаряд. Взрыватель. Заряд."
         Height          =   2800
         Left            =   8700
         TabIndex        =   1
         Top             =   7200
         Width           =   7400
         Begin VB.ComboBox psnar1 
            Height          =   405
            ItemData        =   "Form3.frx":00FE
            Left            =   1300
            List            =   "Form3.frx":010E
            TabIndex        =   13
            Text            =   "ОФ"
            Top             =   900
            Width           =   1200
         End
         Begin VB.ComboBox psnar2 
            Height          =   405
            ItemData        =   "Form3.frx":0122
            Left            =   2700
            List            =   "Form3.frx":0132
            TabIndex        =   12
            Text            =   "ОФ"
            Top             =   900
            Width           =   1200
         End
         Begin VB.ComboBox psnar3 
            Height          =   405
            ItemData        =   "Form3.frx":0146
            Left            =   4100
            List            =   "Form3.frx":0156
            TabIndex        =   11
            Text            =   "ОФ"
            Top             =   900
            Width           =   1200
         End
         Begin VB.ComboBox psnar4 
            Height          =   405
            ItemData        =   "Form3.frx":016A
            Left            =   5500
            List            =   "Form3.frx":017A
            TabIndex        =   10
            Text            =   "ОФ"
            Top             =   900
            Width           =   1200
         End
         Begin VB.ComboBox pvzr1 
            Height          =   405
            ItemData        =   "Form3.frx":018E
            Left            =   1300
            List            =   "Form3.frx":01A1
            TabIndex        =   9
            Text            =   "РГМ"
            Top             =   1400
            Width           =   1200
         End
         Begin VB.ComboBox pvzr2 
            Height          =   405
            ItemData        =   "Form3.frx":01C0
            Left            =   2700
            List            =   "Form3.frx":01D3
            TabIndex        =   8
            Text            =   "РГМ"
            Top             =   1400
            Width           =   1200
         End
         Begin VB.ComboBox pvzr3 
            Height          =   405
            ItemData        =   "Form3.frx":01F2
            Left            =   4100
            List            =   "Form3.frx":0205
            TabIndex        =   7
            Text            =   "РГМ"
            Top             =   1400
            Width           =   1200
         End
         Begin VB.ComboBox pvzr4 
            Height          =   405
            ItemData        =   "Form3.frx":0224
            Left            =   5520
            List            =   "Form3.frx":0237
            TabIndex        =   6
            Text            =   "РГМ"
            Top             =   1400
            Width           =   1200
         End
         Begin VB.ComboBox pzar1 
            Height          =   405
            ItemData        =   "Form3.frx":0256
            Left            =   1300
            List            =   "Form3.frx":026C
            TabIndex        =   5
            Text            =   "Полн"
            Top             =   1900
            Width           =   1200
         End
         Begin VB.ComboBox pzar2 
            Height          =   405
            ItemData        =   "Form3.frx":0294
            Left            =   2700
            List            =   "Form3.frx":02AA
            TabIndex        =   4
            Text            =   "Полн"
            Top             =   1900
            Width           =   1200
         End
         Begin VB.ComboBox pzar3 
            Height          =   405
            ItemData        =   "Form3.frx":02D2
            Left            =   4100
            List            =   "Form3.frx":02E8
            TabIndex        =   3
            Text            =   "Полн"
            Top             =   1900
            Width           =   1200
         End
         Begin VB.ComboBox pzar4 
            Height          =   405
            ItemData        =   "Form3.frx":0310
            Left            =   5500
            List            =   "Form3.frx":0326
            TabIndex        =   2
            Text            =   "Полн"
            Top             =   1900
            Width           =   1200
         End
         Begin VB.Label Label61 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1 Бат"
            Height          =   300
            Left            =   1300
            TabIndex        =   20
            Top             =   400
            Width           =   600
         End
         Begin VB.Label Label62 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2 Бат"
            Height          =   300
            Left            =   2700
            TabIndex        =   19
            Top             =   400
            Width           =   600
         End
         Begin VB.Label Label63 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3 Бат"
            Height          =   300
            Left            =   4100
            TabIndex        =   18
            Top             =   400
            Width           =   600
         End
         Begin VB.Label Label64 
            BackColor       =   &H00C0C0C0&
            Caption         =   "4 Бат"
            Height          =   300
            Left            =   5500
            TabIndex        =   17
            Top             =   400
            Width           =   600
         End
         Begin VB.Label Label65 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Снаряд"
            Height          =   300
            Left            =   300
            TabIndex        =   16
            Top             =   900
            Width           =   900
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Взрыв."
            Height          =   300
            Left            =   300
            TabIndex        =   15
            Top             =   1400
            Width           =   800
         End
         Begin VB.Label Label67 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Заряд"
            Height          =   300
            Left            =   300
            TabIndex        =   14
            Top             =   1900
            Width           =   800
         End
      End
   End
End
Attribute VB_Name = "YO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
