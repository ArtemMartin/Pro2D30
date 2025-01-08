VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.OCX"
Begin VB.Form SIYO 
   BackColor       =   &H00808080&
   Caption         =   "—Ë”Œ ‰Ë‚ 2—1"
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
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   19394
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   706
      BackColor       =   4210688
      ForeColor       =   16384
      TabCaption(0)   =   "¡Œ≈¬Œ… œŒ–ﬂƒŒ "
      TabPicture(0)   =   "SIYOdiv2C1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Œ√Õ≈¬¿ﬂ «¿ƒ¿◊¿"
      TabPicture(1)   =   "SIYOdiv2C1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "œ–»—“–≈À ¿"
      TabPicture(2)   =   "SIYOdiv2C1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "–≈œ≈–"
      TabPicture(3)   =   "SIYOdiv2C1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "«¿√–¿ƒ Œ√ŒÕ‹"
      TabPicture(4)   =   "SIYOdiv2C1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Height          =   10200
         Left            =   100
         TabIndex        =   171
         Top             =   500
         Width           =   11535
         Begin TabDlg.SSTab SSTab2 
            Height          =   9735
            Left            =   105
            TabIndex        =   172
            Top             =   105
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   17171
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   706
            BackColor       =   128
            ForeColor       =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
               Size            =   14.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "÷ÂÎ¸ ‰Îˇ ‰Ë‚ËÁËÓÌ‡"
            TabPicture(0)   =   "SIYOdiv2C1.frx":008C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame7"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "–‡ÔÂ‰ÂÎÂÌËÂ"
            TabPicture(1)   =   "SIYOdiv2C1.frx":00A8
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Frame8"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame8 
               BackColor       =   &H00C0C0C0&
               Height          =   9000
               Left            =   100
               TabIndex        =   213
               Top             =   500
               Width           =   10935
               Begin VB.ComboBox Combo14 
                  Height          =   405
                  Left            =   4920
                  TabIndex        =   253
                  Text            =   "Combo14"
                  Top             =   3840
                  Width           =   855
               End
               Begin VB.CommandButton Command7 
                  Caption         =   "Command7"
                  Height          =   900
                  Left            =   2500
                  TabIndex        =   248
                  Top             =   4200
                  Width           =   1100
               End
               Begin VB.ComboBox Combo13 
                  Height          =   405
                  Left            =   800
                  TabIndex        =   247
                  Text            =   "Combo13"
                  Top             =   3700
                  Width           =   855
               End
               Begin VB.TextBox Text110 
                  Height          =   405
                  Left            =   800
                  TabIndex        =   246
                  Text            =   "Text110"
                  Top             =   5200
                  Width           =   1000
               End
               Begin VB.TextBox Text109 
                  Height          =   405
                  Left            =   800
                  TabIndex        =   245
                  Text            =   "Text109"
                  Top             =   4700
                  Width           =   1500
               End
               Begin VB.TextBox Text108 
                  Height          =   405
                  Left            =   800
                  TabIndex        =   244
                  Text            =   "Text108"
                  Top             =   4200
                  Width           =   1500
               End
               Begin VB.CommandButton Command6 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "–≈ÿ»“‹"
                  Height          =   900
                  Left            =   9600
                  Style           =   1  'Graphical
                  TabIndex        =   238
                  Top             =   1500
                  Width           =   1100
               End
               Begin VB.TextBox Text107 
                  Height          =   400
                  Left            =   7900
                  TabIndex        =   237
                  Text            =   "0"
                  Top             =   2500
                  Width           =   1000
               End
               Begin VB.TextBox Text106 
                  Height          =   400
                  Left            =   7900
                  TabIndex        =   236
                  Text            =   "0"
                  Top             =   2000
                  Width           =   1500
               End
               Begin VB.TextBox Text105 
                  Height          =   400
                  Left            =   7900
                  TabIndex        =   235
                  Text            =   "0"
                  Top             =   1500
                  Width           =   1500
               End
               Begin VB.CommandButton Command5 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "–≈ÿ»“‹"
                  Height          =   900
                  Left            =   6100
                  Style           =   1  'Graphical
                  TabIndex        =   231
                  Top             =   1500
                  Width           =   1100
               End
               Begin VB.TextBox Text104 
                  Height          =   405
                  Left            =   4400
                  TabIndex        =   230
                  Text            =   "0"
                  Top             =   2500
                  Width           =   1000
               End
               Begin VB.TextBox Text103 
                  Height          =   405
                  Left            =   4400
                  TabIndex        =   229
                  Text            =   "0"
                  Top             =   2000
                  Width           =   1500
               End
               Begin VB.TextBox Text102 
                  Height          =   405
                  Left            =   4400
                  TabIndex        =   228
                  Text            =   "0"
                  Top             =   1500
                  Width           =   1500
               End
               Begin VB.CommandButton Command4 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "–≈ÿ»“‹"
                  Height          =   900
                  Left            =   2500
                  Style           =   1  'Graphical
                  TabIndex        =   227
                  Top             =   1500
                  Width           =   1100
               End
               Begin VB.TextBox Text101 
                  Height          =   405
                  Left            =   800
                  TabIndex        =   223
                  Text            =   "0"
                  Top             =   2500
                  Width           =   1000
               End
               Begin VB.TextBox Text100 
                  Height          =   405
                  Left            =   800
                  TabIndex        =   222
                  Text            =   "0"
                  Top             =   2000
                  Width           =   1500
               End
               Begin VB.TextBox Text99 
                  Height          =   405
                  Left            =   800
                  TabIndex        =   221
                  Text            =   "0"
                  Top             =   1500
                  Width           =   1500
               End
               Begin VB.Label Label111 
                  Caption         =   "Label111"
                  Height          =   300
                  Left            =   4080
                  TabIndex        =   252
                  Top             =   5400
                  Width           =   495
               End
               Begin VB.Label Label110 
                  Caption         =   "Label110"
                  Height          =   300
                  Left            =   4080
                  TabIndex        =   251
                  Top             =   4800
                  Width           =   375
               End
               Begin VB.Label Label109 
                  Caption         =   "Label109"
                  Height          =   300
                  Left            =   3960
                  TabIndex        =   250
                  Top             =   4320
                  Width           =   495
               End
               Begin VB.Label Label108 
                  Caption         =   "Label108"
                  Height          =   300
                  Left            =   3960
                  TabIndex        =   249
                  Top             =   3720
                  Width           =   495
               End
               Begin VB.Label Label107 
                  Caption         =   "Label107"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   243
                  Top             =   5200
                  Width           =   500
               End
               Begin VB.Label Label106 
                  Caption         =   "Label106"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   242
                  Top             =   4700
                  Width           =   500
               End
               Begin VB.Label Label105 
                  Caption         =   "Label105"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   241
                  Top             =   4200
                  Width           =   500
               End
               Begin VB.Label Label104 
                  Caption         =   "π Õœ"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   240
                  Top             =   3700
                  Width           =   800
               End
               Begin VB.Label Label103 
                  Caption         =   "Label103"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   239
                  Top             =   3200
                  Width           =   3615
               End
               Begin VB.Label Label102 
                  Caption         =   "h="
                  Height          =   300
                  Left            =   7300
                  TabIndex        =   234
                  Top             =   2500
                  Width           =   500
               End
               Begin VB.Label Label101 
                  Caption         =   "ƒ="
                  Height          =   300
                  Left            =   7300
                  TabIndex        =   233
                  Top             =   2000
                  Width           =   500
               End
               Begin VB.Label Label100 
                  Caption         =   "¿="
                  Height          =   300
                  Left            =   7300
                  TabIndex        =   232
                  Top             =   1500
                  Width           =   500
               End
               Begin VB.Label Label99 
                  Caption         =   "h="
                  Height          =   300
                  Left            =   3700
                  TabIndex        =   226
                  Top             =   2500
                  Width           =   500
               End
               Begin VB.Label Label98 
                  Caption         =   "ƒ="
                  Height          =   300
                  Left            =   3700
                  TabIndex        =   225
                  Top             =   2000
                  Width           =   500
               End
               Begin VB.Label Label97 
                  Caption         =   "¿="
                  Height          =   300
                  Left            =   3700
                  TabIndex        =   224
                  Top             =   1500
                  Width           =   500
               End
               Begin VB.Label Label96 
                  Caption         =   "h="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   220
                  Top             =   2500
                  Width           =   500
               End
               Begin VB.Label Label95 
                  Caption         =   "”="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   219
                  Top             =   2000
                  Width           =   500
               End
               Begin VB.Label Label92 
                  Caption         =   "’="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   218
                  Top             =   1500
                  Width           =   500
               End
               Begin VB.Label Label94 
                  Caption         =   "             3 ¡‡Ú"
                  Height          =   300
                  Left            =   7300
                  TabIndex        =   217
                  Top             =   400
                  Width           =   2200
               End
               Begin VB.Label Label93 
                  Caption         =   "             2 ¡‡Ú"
                  Height          =   300
                  Left            =   3700
                  TabIndex        =   216
                  Top             =   400
                  Width           =   2200
               End
               Begin VB.Label Label91 
                  Caption         =   "                                                                                   ’, ”"
                  Height          =   300
                  Left            =   105
                  TabIndex        =   215
                  Top             =   1005
                  Width           =   10600
               End
               Begin VB.Label Label90 
                  Caption         =   "             1 ¡‡Ú"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   214
                  Top             =   400
                  Width           =   2200
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00C0C0C0&
               Height          =   9000
               Left            =   -74900
               TabIndex        =   173
               Top             =   500
               Width           =   10695
               Begin VB.CommandButton Command3 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "–≈ÿ»“‹"
                  Height          =   900
                  Left            =   6000
                  Style           =   1  'Graphical
                  TabIndex        =   212
                  Top             =   6200
                  Width           =   1200
               End
               Begin VB.TextBox Text98 
                  Height          =   405
                  Left            =   4100
                  TabIndex        =   211
                  Text            =   "0"
                  Top             =   7200
                  Width           =   1000
               End
               Begin VB.TextBox Text97 
                  Height          =   405
                  Left            =   4100
                  TabIndex        =   210
                  Text            =   "0"
                  Top             =   6700
                  Width           =   1500
               End
               Begin VB.TextBox Text96 
                  Height          =   405
                  Left            =   1200
                  TabIndex        =   209
                  Text            =   "0"
                  Top             =   7200
                  Width           =   1000
               End
               Begin VB.TextBox Text95 
                  Height          =   405
                  Left            =   1200
                  TabIndex        =   208
                  Text            =   "0"
                  Top             =   6700
                  Width           =   1500
               End
               Begin VB.ComboBox Combo12 
                  Height          =   405
                  ItemData        =   "SIYOdiv2C1.frx":00C4
                  Left            =   4100
                  List            =   "SIYOdiv2C1.frx":00D7
                  TabIndex        =   207
                  Text            =   "1"
                  Top             =   6200
                  Width           =   800
               End
               Begin VB.ComboBox Combo11 
                  Height          =   405
                  ItemData        =   "SIYOdiv2C1.frx":00EA
                  Left            =   1200
                  List            =   "SIYOdiv2C1.frx":00FD
                  TabIndex        =   206
                  Text            =   "1"
                  Top             =   6200
                  Width           =   800
               End
               Begin VB.CommandButton Command2 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "–≈ÿ»“‹"
                  Height          =   900
                  Left            =   3100
                  Style           =   1  'Graphical
                  TabIndex        =   197
                  Top             =   3600
                  Width           =   1200
               End
               Begin VB.TextBox Text94 
                  Height          =   405
                  Left            =   1200
                  TabIndex        =   195
                  Text            =   "0"
                  Top             =   4600
                  Width           =   1000
               End
               Begin VB.TextBox Text93 
                  Height          =   405
                  Left            =   1200
                  TabIndex        =   194
                  Text            =   "0"
                  Top             =   4100
                  Width           =   1500
               End
               Begin VB.TextBox Text92 
                  Height          =   405
                  Left            =   1200
                  TabIndex        =   193
                  Text            =   "0"
                  Top             =   3600
                  Width           =   1500
               End
               Begin VB.ComboBox Combo10 
                  Height          =   405
                  ItemData        =   "SIYOdiv2C1.frx":0110
                  Left            =   1200
                  List            =   "SIYOdiv2C1.frx":0123
                  TabIndex        =   192
                  Text            =   "1"
                  Top             =   3100
                  Width           =   800
               End
               Begin VB.CommandButton Command1 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "–≈ÿ»“‹"
                  Height          =   900
                  Left            =   2640
                  Style           =   1  'Graphical
                  TabIndex        =   186
                  Top             =   900
                  Width           =   1200
               End
               Begin VB.TextBox Text91 
                  Height          =   405
                  Left            =   5100
                  TabIndex        =   185
                  Text            =   "0"
                  Top             =   1400
                  Width           =   1000
               End
               Begin VB.TextBox Text90 
                  Height          =   405
                  Left            =   5100
                  TabIndex        =   184
                  Text            =   "0"
                  Top             =   900
                  Width           =   1000
               End
               Begin VB.TextBox Text89 
                  Height          =   405
                  Left            =   800
                  TabIndex        =   180
                  Text            =   "0"
                  Top             =   1900
                  Width           =   1000
               End
               Begin VB.TextBox Text88 
                  Height          =   400
                  Left            =   800
                  TabIndex        =   179
                  Text            =   "0"
                  Top             =   1400
                  Width           =   1500
               End
               Begin VB.TextBox Text87 
                  Height          =   400
                  Left            =   800
                  TabIndex        =   178
                  Text            =   "0"
                  Top             =   900
                  Width           =   1500
               End
               Begin VB.Label Label89 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Ãˆ="
                  Height          =   300
                  Left            =   3000
                  TabIndex        =   205
                  Top             =   7200
                  Width           =   500
               End
               Begin VB.Label Label88 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "¿="
                  Height          =   300
                  Left            =   3000
                  TabIndex        =   204
                  Top             =   6700
                  Width           =   500
               End
               Begin VB.Label Label87 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "π  Õœ"
                  Height          =   300
                  Left            =   3000
                  TabIndex        =   203
                  Top             =   6200
                  Width           =   1000
               End
               Begin VB.Label Label86 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Ãˆ="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   202
                  Top             =   7200
                  Width           =   500
               End
               Begin VB.Label Label85 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "¿="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   201
                  Top             =   6700
                  Width           =   500
               End
               Begin VB.Label Label84 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "π  Õœ"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   200
                  Top             =   6200
                  Width           =   1000
               End
               Begin VB.Label Label83 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "      œ–¿¬€…"
                  Height          =   300
                  Left            =   3000
                  TabIndex        =   199
                  Top             =   5700
                  Width           =   1935
               End
               Begin VB.Label Label82 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "     À≈¬€…"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   198
                  Top             =   5700
                  Width           =   1575
               End
               Begin VB.Label Label81 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "                               —Œœ–ﬂ∆≈Õ ¿"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   196
                  Top             =   5200
                  Width           =   5500
               End
               Begin VB.Label Label80 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Ãˆ="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   191
                  Top             =   4600
                  Width           =   500
               End
               Begin VB.Label Label79 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "ƒ="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   190
                  Top             =   4100
                  Width           =   500
               End
               Begin VB.Label Label78 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "¿="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   189
                  Top             =   3600
                  Width           =   500
               End
               Begin VB.Label Label77 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "π  Õœ "
                  Height          =   300
                  Left            =   100
                  TabIndex        =   188
                  Top             =   3100
                  Width           =   1000
               End
               Begin VB.Label Label76 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "                    ¿, ƒ"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   187
                  Top             =   2600
                  Width           =   2600
               End
               Begin VB.Label Label75 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "√ÎÛ·ËÌ‡"
                  Height          =   300
                  Left            =   4000
                  TabIndex        =   183
                  Top             =   1400
                  Width           =   1000
               End
               Begin VB.Label Label74 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "‘ÓÌÚ"
                  Height          =   300
                  Left            =   4000
                  TabIndex        =   182
                  Top             =   900
                  Width           =   1000
               End
               Begin VB.Label Label73 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "    –‡ÁÏÂ˚ ˆÂÎË"
                  Height          =   300
                  Left            =   4000
                  TabIndex        =   181
                  Top             =   400
                  Width           =   2175
               End
               Begin VB.Label Label72 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "h="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   177
                  Top             =   1900
                  Width           =   500
               End
               Begin VB.Label Label71 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "”="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   176
                  Top             =   1400
                  Width           =   500
               End
               Begin VB.Label Label70 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "’="
                  Height          =   300
                  Left            =   100
                  TabIndex        =   175
                  Top             =   900
                  Width           =   500
               End
               Begin VB.Label Label69 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "               ’, ”"
                  Height          =   300
                  Left            =   100
                  TabIndex        =   174
                  Top             =   400
                  Width           =   2200
               End
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¬˚·Ó —Ì‡ˇ‰, ¬Á˚‚‡ÚÂÎ¸, «‡ˇ‰."
         Height          =   3615
         Left            =   -74900
         TabIndex        =   155
         Top             =   6900
         Width           =   7815
         Begin VB.ComboBox Combo9 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":0136
            Left            =   4200
            List            =   "SIYOdiv2C1.frx":014C
            TabIndex        =   170
            Text            =   "œŒÀÕ"
            Top             =   2200
            Width           =   1200
         End
         Begin VB.ComboBox Combo8 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":0178
            Left            =   2900
            List            =   "SIYOdiv2C1.frx":018E
            TabIndex        =   169
            Text            =   "œŒÀÕ"
            Top             =   2200
            Width           =   1200
         End
         Begin VB.ComboBox Combo7 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":01BA
            Left            =   1500
            List            =   "SIYOdiv2C1.frx":01D0
            TabIndex        =   168
            Text            =   "œŒÀÕ"
            Top             =   2200
            Width           =   1200
         End
         Begin VB.ComboBox Combo6 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":01FC
            Left            =   4200
            List            =   "SIYOdiv2C1.frx":020F
            TabIndex        =   167
            Text            =   "–√Ã"
            Top             =   1600
            Width           =   1200
         End
         Begin VB.ComboBox Combo5 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":0232
            Left            =   2900
            List            =   "SIYOdiv2C1.frx":0245
            TabIndex        =   166
            Text            =   "–√Ã"
            Top             =   1600
            Width           =   1200
         End
         Begin VB.ComboBox Combo4 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":0268
            Left            =   1500
            List            =   "SIYOdiv2C1.frx":027B
            TabIndex        =   165
            Text            =   "–√Ã"
            Top             =   1600
            Width           =   1200
         End
         Begin VB.ComboBox Combo3 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":029E
            Left            =   4200
            List            =   "SIYOdiv2C1.frx":02AE
            TabIndex        =   164
            Text            =   "Œ‘"
            Top             =   1000
            Width           =   1200
         End
         Begin VB.ComboBox Combo2 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":02C2
            Left            =   2900
            List            =   "SIYOdiv2C1.frx":02D2
            TabIndex        =   163
            Text            =   "Œ‘"
            Top             =   1000
            Width           =   1200
         End
         Begin VB.ComboBox Combo1 
            Height          =   405
            ItemData        =   "SIYOdiv2C1.frx":02E6
            Left            =   1500
            List            =   "SIYOdiv2C1.frx":02F6
            TabIndex        =   162
            Text            =   "Œ‘"
            Top             =   1000
            Width           =   1200
         End
         Begin VB.Label Label68 
            BackColor       =   &H00C0C0C0&
            Caption         =   "«‡ˇ‰"
            Height          =   300
            Left            =   100
            TabIndex        =   161
            Top             =   2200
            Width           =   1200
         End
         Begin VB.Label Label67 
            BackColor       =   &H00C0C0C0&
            Caption         =   "¬Á˚‚‡ÚÂÎ¸"
            Height          =   300
            Left            =   100
            TabIndex        =   160
            Top             =   1600
            Width           =   1400
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C0C0&
            Caption         =   "—Ì‡ˇ‰"
            Height          =   300
            Left            =   100
            TabIndex        =   159
            Top             =   1000
            Width           =   1200
         End
         Begin VB.Label Label65 
            BackColor       =   &H00C0C0C0&
            Caption         =   "   3 ¡‡Ú"
            Height          =   300
            Left            =   4200
            TabIndex        =   158
            Top             =   400
            Width           =   1000
         End
         Begin VB.Label Label64 
            BackColor       =   &H00C0C0C0&
            Caption         =   "   2 ¡‡Ú"
            Height          =   300
            Left            =   2900
            TabIndex        =   157
            Top             =   400
            Width           =   1000
         End
         Begin VB.Label Label63 
            BackColor       =   &H00C0C0C0&
            Caption         =   "   1 ¡‡Ú"
            Height          =   300
            Left            =   1500
            TabIndex        =   156
            Top             =   400
            Width           =   1000
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÃÂÚÂÓ‰‡ÌÌ˚Â"
         Height          =   6300
         Left            =   -61400
         TabIndex        =   91
         Top             =   500
         Width           =   6495
         Begin VB.TextBox Text86 
            Height          =   405
            Left            =   5200
            TabIndex        =   154
            Text            =   "0"
            Top             =   4900
            Width           =   700
         End
         Begin VB.TextBox Text85 
            Height          =   405
            Left            =   4400
            TabIndex        =   153
            Text            =   "0"
            Top             =   4900
            Width           =   700
         End
         Begin VB.TextBox Text84 
            Height          =   405
            Left            =   3600
            TabIndex        =   152
            Text            =   "0"
            Top             =   4900
            Width           =   700
         End
         Begin VB.TextBox Text83 
            Height          =   405
            Left            =   5200
            TabIndex        =   151
            Text            =   "0"
            Top             =   4500
            Width           =   700
         End
         Begin VB.TextBox Text82 
            Height          =   405
            Left            =   4400
            TabIndex        =   150
            Text            =   "0"
            Top             =   4500
            Width           =   700
         End
         Begin VB.TextBox Text81 
            Height          =   405
            Left            =   3600
            TabIndex        =   149
            Text            =   "0"
            Top             =   4500
            Width           =   700
         End
         Begin VB.TextBox Text80 
            Height          =   405
            Left            =   5200
            TabIndex        =   148
            Text            =   "0"
            Top             =   4100
            Width           =   700
         End
         Begin VB.TextBox Text79 
            Height          =   405
            Left            =   4400
            TabIndex        =   147
            Text            =   "0"
            Top             =   4100
            Width           =   700
         End
         Begin VB.TextBox Text78 
            Height          =   405
            Left            =   3600
            TabIndex        =   146
            Text            =   "0"
            Top             =   4100
            Width           =   700
         End
         Begin VB.TextBox Text77 
            Height          =   405
            Left            =   5200
            TabIndex        =   145
            Text            =   "0"
            Top             =   3700
            Width           =   700
         End
         Begin VB.TextBox Text76 
            Height          =   405
            Left            =   4400
            TabIndex        =   144
            Text            =   "0"
            Top             =   3700
            Width           =   700
         End
         Begin VB.TextBox Text75 
            Height          =   405
            Left            =   3600
            TabIndex        =   143
            Text            =   "0"
            Top             =   3700
            Width           =   700
         End
         Begin VB.TextBox Text74 
            Height          =   405
            Left            =   5200
            TabIndex        =   142
            Text            =   "0"
            Top             =   3300
            Width           =   700
         End
         Begin VB.TextBox Text73 
            Height          =   405
            Left            =   4400
            TabIndex        =   141
            Text            =   "0"
            Top             =   3300
            Width           =   700
         End
         Begin VB.TextBox Text72 
            Height          =   405
            Left            =   3600
            TabIndex        =   140
            Text            =   "0"
            Top             =   3300
            Width           =   700
         End
         Begin VB.TextBox Text71 
            Height          =   405
            Left            =   5200
            TabIndex        =   139
            Text            =   "0"
            Top             =   2900
            Width           =   700
         End
         Begin VB.TextBox Text70 
            Height          =   405
            Left            =   4400
            TabIndex        =   138
            Text            =   "0"
            Top             =   2900
            Width           =   700
         End
         Begin VB.TextBox Text69 
            Height          =   405
            Left            =   3600
            TabIndex        =   137
            Text            =   "0"
            Top             =   2900
            Width           =   700
         End
         Begin VB.TextBox Text68 
            Height          =   405
            Left            =   5200
            TabIndex        =   136
            Text            =   "0"
            Top             =   2500
            Width           =   700
         End
         Begin VB.TextBox Text67 
            Height          =   405
            Left            =   4400
            TabIndex        =   135
            Text            =   "0"
            Top             =   2500
            Width           =   700
         End
         Begin VB.TextBox Text66 
            Height          =   405
            Left            =   3600
            TabIndex        =   134
            Text            =   "0"
            Top             =   2500
            Width           =   700
         End
         Begin VB.TextBox Text65 
            Height          =   405
            Left            =   5200
            TabIndex        =   133
            Text            =   "0"
            Top             =   2100
            Width           =   700
         End
         Begin VB.TextBox Text64 
            Height          =   405
            Left            =   4400
            TabIndex        =   132
            Text            =   "0"
            Top             =   2100
            Width           =   700
         End
         Begin VB.TextBox Text63 
            Height          =   405
            Left            =   3600
            TabIndex        =   131
            Text            =   "0"
            Top             =   2100
            Width           =   700
         End
         Begin VB.TextBox Text62 
            Height          =   405
            Left            =   5200
            TabIndex        =   130
            Text            =   "0"
            Top             =   1700
            Width           =   700
         End
         Begin VB.TextBox Text61 
            Height          =   405
            Left            =   4400
            TabIndex        =   129
            Text            =   "0"
            Top             =   1700
            Width           =   700
         End
         Begin VB.TextBox Text60 
            Height          =   405
            Left            =   3600
            TabIndex        =   128
            Text            =   "0"
            Top             =   1700
            Width           =   700
         End
         Begin VB.TextBox Text59 
            Height          =   405
            Left            =   5200
            TabIndex        =   127
            Text            =   "0"
            Top             =   1300
            Width           =   700
         End
         Begin VB.TextBox Text58 
            Height          =   405
            Left            =   4400
            TabIndex        =   126
            Text            =   "0"
            Top             =   1300
            Width           =   700
         End
         Begin VB.TextBox Text57 
            Height          =   405
            Left            =   3600
            TabIndex        =   125
            Text            =   "0"
            Top             =   1300
            Width           =   700
         End
         Begin VB.TextBox Text56 
            Height          =   405
            Left            =   5200
            TabIndex        =   124
            Text            =   "0"
            Top             =   900
            Width           =   700
         End
         Begin VB.TextBox Text55 
            Height          =   405
            Left            =   4400
            TabIndex        =   123
            Text            =   "0"
            Top             =   900
            Width           =   700
         End
         Begin VB.TextBox Text54 
            Height          =   405
            Left            =   3600
            TabIndex        =   122
            Text            =   "0"
            Top             =   900
            Width           =   700
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            Height          =   1695
            Left            =   100
            TabIndex        =   103
            Top             =   3400
            Width           =   2100
            Begin VB.OptionButton Option2 
               BackColor       =   &H00C0C0C0&
               Height          =   400
               Left            =   1200
               TabIndex        =   107
               Top             =   900
               Width           =   500
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0C0C0&
               Height          =   400
               Left            =   200
               TabIndex        =   106
               Top             =   900
               Value           =   -1  'True
               Width           =   500
            End
            Begin VB.Label Label48 
               BackColor       =   &H00C0C0C0&
               Caption         =   "¬–-2"
               Height          =   300
               Left            =   1200
               TabIndex        =   105
               Top             =   300
               Width           =   600
            End
            Begin VB.Label Label47 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ƒÃ "
               Height          =   300
               Left            =   100
               TabIndex        =   104
               Top             =   300
               Width           =   600
            End
         End
         Begin VB.TextBox Text53 
            Height          =   405
            Left            =   1200
            TabIndex        =   102
            Text            =   "0"
            Top             =   2900
            Width           =   1000
         End
         Begin VB.TextBox Text52 
            Height          =   405
            Left            =   1200
            TabIndex        =   101
            Text            =   "0"
            Top             =   2400
            Width           =   1000
         End
         Begin VB.TextBox Text51 
            Height          =   405
            Left            =   1200
            TabIndex        =   100
            Text            =   "0"
            Top             =   1400
            Width           =   1000
         End
         Begin VB.TextBox Text50 
            Height          =   405
            Left            =   1200
            TabIndex        =   99
            Text            =   "0"
            Top             =   900
            Width           =   1000
         End
         Begin VB.TextBox Text49 
            Height          =   405
            Left            =   1200
            TabIndex        =   98
            Text            =   "0"
            Top             =   400
            Width           =   1000
         End
         Begin VB.Label Label62 
            BackColor       =   &H00C0C0C0&
            Caption         =   "60"
            Height          =   300
            Left            =   2700
            TabIndex        =   121
            Top             =   4900
            Width           =   500
         End
         Begin VB.Label Label61 
            BackColor       =   &H00C0C0C0&
            Caption         =   "50"
            Height          =   300
            Left            =   2700
            TabIndex        =   120
            Top             =   4500
            Width           =   500
         End
         Begin VB.Label Label60 
            BackColor       =   &H00C0C0C0&
            Caption         =   "40"
            Height          =   300
            Left            =   2700
            TabIndex        =   119
            Top             =   4100
            Width           =   500
         End
         Begin VB.Label Label59 
            BackColor       =   &H00C0C0C0&
            Caption         =   "30"
            Height          =   300
            Left            =   2700
            TabIndex        =   118
            Top             =   3700
            Width           =   500
         End
         Begin VB.Label Label58 
            BackColor       =   &H00C0C0C0&
            Caption         =   "24"
            Height          =   300
            Left            =   2700
            TabIndex        =   117
            Top             =   3300
            Width           =   500
         End
         Begin VB.Label Label57 
            BackColor       =   &H00C0C0C0&
            Caption         =   "20"
            Height          =   300
            Left            =   2700
            TabIndex        =   116
            Top             =   2900
            Width           =   500
         End
         Begin VB.Label Label56 
            BackColor       =   &H00C0C0C0&
            Caption         =   "16"
            Height          =   300
            Left            =   2700
            TabIndex        =   115
            Top             =   2500
            Width           =   500
         End
         Begin VB.Label Label55 
            BackColor       =   &H00C0C0C0&
            Caption         =   "12"
            Height          =   300
            Left            =   2700
            TabIndex        =   114
            Top             =   2100
            Width           =   500
         End
         Begin VB.Label Label54 
            BackColor       =   &H00C0C0C0&
            Caption         =   "08"
            Height          =   300
            Left            =   2700
            TabIndex        =   113
            Top             =   1700
            Width           =   500
         End
         Begin VB.Label Label53 
            BackColor       =   &H00C0C0C0&
            Caption         =   "04"
            Height          =   300
            Left            =   2700
            TabIndex        =   112
            Top             =   1300
            Width           =   500
         End
         Begin VB.Label Label52 
            BackColor       =   &H00C0C0C0&
            Caption         =   "02"
            Height          =   300
            Left            =   2700
            TabIndex        =   111
            Top             =   900
            Width           =   500
         End
         Begin VB.Label Label51 
            BackColor       =   &H00C0C0C0&
            Caption         =   "  W"
            Height          =   300
            Left            =   5300
            TabIndex        =   110
            Top             =   400
            Width           =   500
         End
         Begin VB.Label Label50 
            BackColor       =   &H00C0C0C0&
            Caption         =   " Aw"
            Height          =   300
            Left            =   4500
            TabIndex        =   109
            Top             =   400
            Width           =   500
         End
         Begin VB.Label Label49 
            BackColor       =   &H00C0C0C0&
            Caption         =   " dT"
            Height          =   300
            Left            =   3700
            TabIndex        =   108
            Top             =   400
            Width           =   500
         End
         Begin VB.Label Label46 
            BackColor       =   &H00C0C0C0&
            Caption         =   "W="
            Height          =   300
            Left            =   100
            TabIndex        =   97
            Top             =   2900
            Width           =   1000
         End
         Begin VB.Label Label45 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Aw="
            Height          =   300
            Left            =   100
            TabIndex        =   96
            Top             =   2400
            Width           =   1000
         End
         Begin VB.Label Label44 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”„ÓÎ ‚ÂÚ‡ ‚ ƒ”"
            Height          =   255
            Left            =   105
            TabIndex        =   95
            Top             =   1905
            Width           =   2100
         End
         Begin VB.Label Label43 
            BackColor       =   &H00C0C0C0&
            Caption         =   "“‚="
            Height          =   300
            Left            =   100
            TabIndex        =   94
            Top             =   1400
            Width           =   1000
         End
         Begin VB.Label Label42 
            BackColor       =   &H00C0C0C0&
            Caption         =   "H="
            Height          =   300
            Left            =   100
            TabIndex        =   93
            Top             =   900
            Width           =   1000
         End
         Begin VB.Label Label41 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h ÃÂÚÂÓ="
            Height          =   300
            Left            =   100
            TabIndex        =   92
            Top             =   400
            Width           =   1100
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¡‡ÎËÒÚËÍ‡"
         Height          =   6300
         Left            =   -65800
         TabIndex        =   58
         Top             =   500
         Width           =   4215
         Begin VB.TextBox Text48 
            Height          =   405
            Left            =   3100
            TabIndex        =   90
            Text            =   "0"
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox Text47 
            Height          =   405
            Left            =   3100
            TabIndex        =   89
            Text            =   "0"
            Top             =   4300
            Width           =   600
         End
         Begin VB.TextBox Text46 
            Height          =   405
            Left            =   3100
            TabIndex        =   88
            Text            =   "0"
            Top             =   3800
            Width           =   600
         End
         Begin VB.TextBox Text45 
            Height          =   405
            Left            =   3100
            TabIndex        =   87
            Text            =   "0"
            Top             =   3300
            Width           =   600
         End
         Begin VB.TextBox Text44 
            Height          =   405
            Left            =   3100
            TabIndex        =   86
            Text            =   "0"
            Top             =   2800
            Width           =   600
         End
         Begin VB.TextBox Text43 
            Height          =   405
            Left            =   3100
            TabIndex        =   85
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox Text42 
            Height          =   405
            Left            =   2200
            TabIndex        =   84
            Text            =   "0"
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox Text41 
            Height          =   405
            Left            =   2200
            TabIndex        =   83
            Text            =   "0"
            Top             =   4300
            Width           =   600
         End
         Begin VB.TextBox Text40 
            Height          =   405
            Left            =   2200
            TabIndex        =   82
            Text            =   "0"
            Top             =   3800
            Width           =   600
         End
         Begin VB.TextBox Text39 
            Height          =   405
            Left            =   2200
            TabIndex        =   81
            Text            =   "0"
            Top             =   3300
            Width           =   600
         End
         Begin VB.TextBox Text38 
            Height          =   405
            Left            =   2200
            TabIndex        =   80
            Text            =   "0"
            Top             =   2800
            Width           =   600
         End
         Begin VB.TextBox Text37 
            Height          =   405
            Left            =   2200
            TabIndex        =   79
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox Text36 
            Height          =   405
            Left            =   1300
            TabIndex        =   78
            Text            =   "0"
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox Text35 
            Height          =   405
            Left            =   1300
            TabIndex        =   77
            Text            =   "0"
            Top             =   4300
            Width           =   600
         End
         Begin VB.TextBox Text34 
            Height          =   405
            Left            =   1300
            TabIndex        =   76
            Text            =   "0"
            Top             =   3800
            Width           =   600
         End
         Begin VB.TextBox Text33 
            Height          =   405
            Left            =   1300
            TabIndex        =   75
            Text            =   "0"
            Top             =   3300
            Width           =   600
         End
         Begin VB.TextBox Text32 
            Height          =   405
            Left            =   1300
            TabIndex        =   74
            Text            =   "0"
            Top             =   2800
            Width           =   600
         End
         Begin VB.TextBox Text31 
            Height          =   405
            Left            =   1300
            TabIndex        =   73
            Text            =   "0"
            Top             =   2300
            Width           =   600
         End
         Begin VB.TextBox Text30 
            Height          =   405
            Left            =   3100
            TabIndex        =   65
            Text            =   "0"
            Top             =   900
            Width           =   600
         End
         Begin VB.TextBox Text29 
            Height          =   405
            Left            =   2200
            TabIndex        =   64
            Text            =   "0"
            Top             =   900
            Width           =   600
         End
         Begin VB.TextBox Text28 
            Height          =   405
            Left            =   1300
            TabIndex        =   63
            Text            =   "0"
            Top             =   900
            Width           =   600
         End
         Begin VB.Label Label40 
            BackColor       =   &H00C0C0C0&
            Caption         =   "◊ÂÚ‚ÂÚ."
            Height          =   300
            Left            =   100
            TabIndex        =   72
            Top             =   4800
            Width           =   1000
         End
         Begin VB.Label Label39 
            BackColor       =   &H00C0C0C0&
            Caption         =   "“ÂÚ."
            Height          =   300
            Left            =   120
            TabIndex        =   71
            Top             =   4305
            Width           =   1005
         End
         Begin VB.Label Label38 
            BackColor       =   &H00C0C0C0&
            Caption         =   "¬ÚÓ."
            Height          =   300
            Left            =   100
            TabIndex        =   70
            Top             =   3800
            Width           =   1000
         End
         Begin VB.Label Label37 
            BackColor       =   &H00C0C0C0&
            Caption         =   "œÂ‚."
            Height          =   300
            Left            =   100
            TabIndex        =   69
            Top             =   3300
            Width           =   1000
         End
         Begin VB.Label Label36 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”ÏÂÌ¯."
            Height          =   300
            Left            =   100
            TabIndex        =   68
            Top             =   2800
            Width           =   1000
         End
         Begin VB.Label Label35 
            BackColor       =   &H00C0C0C0&
            Caption         =   "œÓÎÌ."
            Height          =   300
            Left            =   100
            TabIndex        =   67
            Top             =   2300
            Width           =   1000
         End
         Begin VB.Label Label34 
            BackColor       =   &H00C0C0C0&
            Caption         =   "œÓÚÂˇ Ì‡˜‡Î¸ÌÓÈ ÒÍÓÓÒÚË %"
            Height          =   300
            Left            =   105
            TabIndex        =   66
            Top             =   1800
            Width           =   3600
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0C0C0&
            Caption         =   "“Á="
            Height          =   300
            Left            =   100
            TabIndex        =   62
            Top             =   900
            Width           =   500
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3 ¡‡Ú"
            Height          =   300
            Left            =   3100
            TabIndex        =   61
            Top             =   400
            Width           =   600
         End
         Begin VB.Label Label31 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2 ¡‡Ú"
            Height          =   300
            Left            =   2200
            TabIndex        =   60
            Top             =   400
            Width           =   600
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1 ¡‡Ú"
            Height          =   300
            Left            =   1300
            TabIndex        =   59
            Top             =   400
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¡ÓÂ‚ÓÈ ÔÓˇ‰ÓÍ"
         Height          =   6300
         Left            =   -74900
         TabIndex        =   1
         Top             =   500
         Width           =   8895
         Begin VB.TextBox Text27 
            Height          =   405
            Left            =   5700
            TabIndex        =   57
            Text            =   "0"
            Top             =   5200
            Width           =   1000
         End
         Begin VB.TextBox Text26 
            Height          =   405
            Left            =   5700
            TabIndex        =   56
            Text            =   "0"
            Top             =   4700
            Width           =   1000
         End
         Begin VB.TextBox Text25 
            Height          =   405
            Left            =   5700
            TabIndex        =   55
            Text            =   "0"
            Top             =   4200
            Width           =   1000
         End
         Begin VB.TextBox Text24 
            Height          =   405
            Left            =   5700
            TabIndex        =   54
            Text            =   "0"
            Top             =   3700
            Width           =   1000
         End
         Begin VB.TextBox Text23 
            Height          =   405
            Left            =   5700
            TabIndex        =   53
            Text            =   "0"
            Top             =   3200
            Width           =   1000
         End
         Begin VB.TextBox Text22 
            Height          =   405
            Left            =   3400
            TabIndex        =   47
            Text            =   "0"
            Top             =   5200
            Width           =   1500
         End
         Begin VB.TextBox Text21 
            Height          =   405
            Left            =   3400
            TabIndex        =   46
            Text            =   "0"
            Top             =   4700
            Width           =   1500
         End
         Begin VB.TextBox Text20 
            Height          =   405
            Left            =   3400
            TabIndex        =   45
            Text            =   "0"
            Top             =   4200
            Width           =   1500
         End
         Begin VB.TextBox Text19 
            Height          =   405
            Left            =   3400
            TabIndex        =   44
            Text            =   "0"
            Top             =   3700
            Width           =   1500
         End
         Begin VB.TextBox Text18 
            Height          =   405
            Left            =   3400
            TabIndex        =   43
            Text            =   "0"
            Top             =   3200
            Width           =   1500
         End
         Begin VB.TextBox Text17 
            Height          =   400
            Left            =   1100
            TabIndex        =   37
            Text            =   "0"
            Top             =   5200
            Width           =   1500
         End
         Begin VB.TextBox Text16 
            Height          =   400
            Left            =   1100
            TabIndex        =   36
            Text            =   "0"
            Top             =   4700
            Width           =   1500
         End
         Begin VB.TextBox Text15 
            Height          =   400
            Left            =   1100
            TabIndex        =   35
            Text            =   "0"
            Top             =   4200
            Width           =   1500
         End
         Begin VB.TextBox Text14 
            Height          =   400
            Left            =   1100
            TabIndex        =   34
            Text            =   "0"
            Top             =   3700
            Width           =   1500
         End
         Begin VB.TextBox Text13 
            Height          =   400
            Left            =   1100
            TabIndex        =   33
            Text            =   "0"
            Top             =   3200
            Width           =   1500
         End
         Begin VB.TextBox Text12 
            Height          =   405
            Left            =   7400
            TabIndex        =   26
            Text            =   "0"
            Top             =   1900
            Width           =   1000
         End
         Begin VB.TextBox Text11 
            Height          =   405
            Left            =   7400
            TabIndex        =   25
            Text            =   "0"
            Top             =   1400
            Width           =   1000
         End
         Begin VB.TextBox Text10 
            Height          =   405
            Left            =   7400
            TabIndex        =   24
            Text            =   "0"
            Top             =   900
            Width           =   1000
         End
         Begin VB.TextBox Text9 
            Height          =   405
            Left            =   5700
            TabIndex        =   20
            Text            =   "0"
            Top             =   1900
            Width           =   1000
         End
         Begin VB.TextBox Text8 
            Height          =   405
            Left            =   5700
            TabIndex        =   19
            Text            =   "0"
            Top             =   1400
            Width           =   1000
         End
         Begin VB.TextBox Text7 
            Height          =   405
            Left            =   5700
            TabIndex        =   18
            Text            =   "0"
            Top             =   900
            Width           =   1000
         End
         Begin VB.TextBox Text6 
            Height          =   405
            Left            =   3400
            TabIndex        =   14
            Text            =   "0"
            Top             =   1900
            Width           =   1500
         End
         Begin VB.TextBox Text5 
            Height          =   405
            Left            =   3400
            TabIndex        =   13
            Text            =   "0"
            Top             =   1400
            Width           =   1500
         End
         Begin VB.TextBox Text4 
            Height          =   405
            Left            =   3400
            TabIndex        =   12
            Text            =   "0"
            Top             =   900
            Width           =   1500
         End
         Begin VB.TextBox Text3 
            Height          =   405
            Left            =   1100
            TabIndex        =   8
            Text            =   "0"
            Top             =   1900
            Width           =   1500
         End
         Begin VB.TextBox Text2 
            Height          =   405
            Left            =   1100
            TabIndex        =   7
            Text            =   "0"
            Top             =   1400
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Height          =   405
            Left            =   1100
            TabIndex        =   6
            Text            =   "0"
            Top             =   900
            Width           =   1500
         End
         Begin VB.Label Label29 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   5100
            TabIndex        =   52
            Top             =   5200
            Width           =   500
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   5100
            TabIndex        =   51
            Top             =   4700
            Width           =   500
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   5100
            TabIndex        =   50
            Top             =   4200
            Width           =   500
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   5100
            TabIndex        =   49
            Top             =   3700
            Width           =   500
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   5100
            TabIndex        =   48
            Top             =   3200
            Width           =   500
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”="
            Height          =   300
            Left            =   2800
            TabIndex        =   42
            Top             =   5200
            Width           =   500
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”="
            Height          =   300
            Left            =   2800
            TabIndex        =   41
            Top             =   4700
            Width           =   500
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”="
            Height          =   300
            Left            =   2800
            TabIndex        =   40
            Top             =   4200
            Width           =   500
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”="
            Height          =   300
            Left            =   2800
            TabIndex        =   39
            Top             =   3700
            Width           =   500
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”="
            Height          =   300
            Left            =   2800
            TabIndex        =   38
            Top             =   3200
            Width           =   500
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Õœ 5 ’="
            Height          =   300
            Left            =   100
            TabIndex        =   32
            Top             =   5200
            Width           =   1000
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Õœ 4 ’="
            Height          =   300
            Left            =   100
            TabIndex        =   31
            Top             =   4700
            Width           =   1000
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Õœ 3 ’="
            Height          =   300
            Left            =   100
            TabIndex        =   30
            Top             =   4200
            Width           =   1000
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Õœ 2 ’="
            Height          =   300
            Left            =   100
            TabIndex        =   29
            Top             =   3700
            Width           =   1000
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Õœ 1 ’="
            Height          =   300
            Left            =   100
            TabIndex        =   28
            Top             =   3200
            Width           =   1000
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "                                                                 Õœ"
            Height          =   300
            Left            =   100
            TabIndex        =   27
            Top             =   2700
            Width           =   8200
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ŒÕ="
            Height          =   300
            Left            =   6840
            TabIndex        =   23
            Top             =   1900
            Width           =   500
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ŒÕ="
            Height          =   300
            Left            =   6840
            TabIndex        =   22
            Top             =   1400
            Width           =   500
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ŒÕ="
            Height          =   300
            Left            =   6840
            TabIndex        =   21
            Top             =   900
            Width           =   500
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   5100
            TabIndex        =   17
            Top             =   1900
            Width           =   500
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   5100
            TabIndex        =   16
            Top             =   1400
            Width           =   500
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "h="
            Height          =   300
            Left            =   5100
            TabIndex        =   15
            Top             =   900
            Width           =   500
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”="
            Height          =   300
            Left            =   2800
            TabIndex        =   11
            Top             =   1900
            Width           =   500
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”="
            Height          =   300
            Left            =   2800
            TabIndex        =   10
            Top             =   1400
            Width           =   500
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "”="
            Height          =   300
            Left            =   2800
            TabIndex        =   9
            Top             =   900
            Width           =   500
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3 ¡‡Ú ’="
            Height          =   300
            Left            =   100
            TabIndex        =   5
            Top             =   1900
            Width           =   1000
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2 ¡‡Ú ’="
            Height          =   300
            Left            =   100
            TabIndex        =   4
            Top             =   1400
            Width           =   1000
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1 ¡‡Ú ’="
            Height          =   300
            Left            =   100
            TabIndex        =   3
            Top             =   900
            Width           =   1000
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "                                                           Œ√Õ≈¬€≈"
            Height          =   255
            Left            =   100
            TabIndex        =   2
            Top             =   400
            Width           =   8300
         End
      End
   End
End
Attribute VB_Name = "SIYO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
