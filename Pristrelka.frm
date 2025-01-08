VERSION 5.00
Begin VB.Form Pristrelka 
   Caption         =   "Ïðèñòðåëêà"
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
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Öåëü"
      Height          =   3255
      Left            =   14160
      TabIndex        =   83
      Top             =   3700
      Width           =   4020
      Begin VB.OptionButton pnzo 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   500
         TabIndex        =   88
         Top             =   2280
         Width           =   375
      End
      Begin VB.OptionButton pkagdomy 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2520
         TabIndex        =   86
         Top             =   1000
         Width           =   615
      End
      Begin VB.OptionButton pvsem 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   500
         TabIndex        =   85
         Top             =   1000
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÍÇÎ"
         Height          =   400
         Left            =   400
         TabIndex        =   89
         Top             =   1500
         Width           =   800
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ðàñïðåäèëåíèå"
         Height          =   405
         Left            =   1800
         TabIndex        =   87
         Top             =   405
         Width           =   2000
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÑÎ"
         Height          =   400
         Left            =   500
         TabIndex        =   84
         Top             =   400
         Width           =   800
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Ðàñ÷åò ñðåäíåãî"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   15500
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   7440
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Âûõîä"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   17760
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   7440
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ñîïðÿæåíêà"
      Height          =   4000
      Left            =   4800
      TabIndex        =   50
      Top             =   3700
      Width           =   9200
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ðåøèòü"
         Height          =   900
         Left            =   6000
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1900
         Width           =   1100
      End
      Begin VB.ComboBox pnbatSopr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":0000
         Left            =   7000
         List            =   "Pristrelka.frx":000D
         TabIndex        =   67
         Text            =   "1"
         Top             =   1400
         Width           =   800
      End
      Begin VB.ComboBox pnkpPrav 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":001A
         Left            =   4100
         List            =   "Pristrelka.frx":002D
         TabIndex        =   65
         Text            =   "1"
         Top             =   1400
         Width           =   800
      End
      Begin VB.TextBox pMrPrav 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4100
         TabIndex        =   64
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pArPrav 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4100
         TabIndex        =   63
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.ComboBox pnkpLev 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":0040
         Left            =   1000
         List            =   "Pristrelka.frx":0053
         TabIndex        =   59
         Text            =   "1"
         Top             =   1400
         Width           =   800
      End
      Begin VB.TextBox pMrLev 
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
         Left            =   1000
         TabIndex        =   58
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pArLev 
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
         Left            =   1000
         TabIndex        =   57
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ Áàò="
         Height          =   300
         Left            =   6000
         TabIndex        =   66
         Top             =   1400
         Width           =   900
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ìöð="
         Height          =   300
         Left            =   3300
         TabIndex        =   62
         Top             =   2400
         Width           =   700
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Àð="
         Height          =   300
         Left            =   3300
         TabIndex        =   61
         Top             =   1900
         Width           =   500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ ÊÏ="
         Height          =   300
         Left            =   3300
         TabIndex        =   60
         Top             =   1400
         Width           =   800
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ìöð="
         Height          =   300
         Left            =   200
         TabIndex        =   56
         Top             =   2400
         Width           =   700
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Àð="
         Height          =   300
         Left            =   200
         TabIndex        =   55
         Top             =   1900
         Width           =   500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ ÊÏ="
         Height          =   300
         Left            =   200
         TabIndex        =   54
         Top             =   1400
         Width           =   800
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "         Ïðàâûé"
         Height          =   300
         Left            =   3300
         TabIndex        =   53
         Top             =   900
         Width           =   2300
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "           Ëåâûé"
         Height          =   300
         Left            =   200
         TabIndex        =   52
         Top             =   900
         Width           =   2300
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "                              ÑÎÏÐßÆÅÍÊÀ"
         Height          =   300
         Left            =   200
         TabIndex        =   51
         Top             =   400
         Width           =   5400
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Êîððåêòóðà"
      Height          =   6800
      Left            =   100
      TabIndex        =   42
      Top             =   3700
      Width           =   4500
      Begin VB.TextBox pvKorTr 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3200
         TabIndex        =   81
         Text            =   "0"
         Top             =   2000
         Width           =   1100
      End
      Begin VB.TextBox pvKorPr 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1100
         TabIndex        =   79
         Text            =   "0"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.TextBox pvPolOP 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2400
         TabIndex        =   77
         Text            =   "0"
         Top             =   4400
         Width           =   1500
      End
      Begin VB.TextBox pvSHy 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1100
         TabIndex        =   76
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvKy 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1100
         TabIndex        =   75
         Text            =   "0"
         Top             =   4400
         Width           =   1000
      End
      Begin VB.TextBox pvPS 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1100
         TabIndex        =   74
         Text            =   "0"
         Top             =   3600
         Width           =   1000
      End
      Begin VB.TextBox pvKorD 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3200
         TabIndex        =   45
         Text            =   "0"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.TextBox pvKorDov 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1100
         TabIndex        =   44
         Text            =   "0"
         Top             =   2000
         Width           =   1100
      End
      Begin VB.TextBox pvKorYr 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1100
         TabIndex        =   43
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Òð="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2400
         TabIndex        =   80
         Top             =   2000
         Width           =   650
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ä="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2400
         TabIndex        =   78
         Top             =   1200
         Width           =   650
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   " ÎÏ"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2800
         TabIndex        =   73
         Top             =   3900
         Width           =   705
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Øó="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   72
         Top             =   5200
         Width           =   700
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Êó="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   71
         Top             =   4400
         Width           =   700
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÏÑ="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   70
         Top             =   3600
         Width           =   700
      End
      Begin VB.Label Label95 
         BackColor       =   &H00C0C0C0&
         Caption         =   "             ÊÎÐÐÅÊÒÓÐÀ"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   195
         TabIndex        =   49
         Top             =   405
         Width           =   4100
      End
      Begin VB.Label Label96 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ïð="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   48
         Top             =   1200
         Width           =   650
      End
      Begin VB.Label Label97 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Äîâ="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   47
         Top             =   2000
         Width           =   800
      End
      Begin VB.Label Label98 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Óð="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   200
         TabIndex        =   46
         Top             =   2800
         Width           =   700
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "dX, dY"
      Height          =   3500
      Index           =   2
      Left            =   14160
      TabIndex        =   33
      Top             =   0
      Width           =   4500
      Begin VB.ComboBox pnbatdXdY 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":0066
         Left            =   1100
         List            =   "Pristrelka.frx":0073
         TabIndex        =   37
         Text            =   "1"
         Top             =   900
         Width           =   1000
      End
      Begin VB.CommandButton PrdXdY 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ðåøèòü"
         Height          =   900
         Index           =   2
         Left            =   2900
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1400
         Width           =   1100
      End
      Begin VB.TextBox pdYr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1100
         TabIndex        =   35
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.TextBox pdXr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1100
         TabIndex        =   34
         Text            =   "0"
         Top             =   1400
         Width           =   1500
      End
      Begin VB.Label Label94 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ Áàò="
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
         Index           =   3
         Left            =   100
         TabIndex        =   41
         Top             =   900
         Width           =   1000
      End
      Begin VB.Label Label92 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dYð="
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
         Left            =   200
         TabIndex        =   40
         Top             =   1900
         Width           =   700
      End
      Begin VB.Label Label91 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dXð="
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
         Left            =   200
         TabIndex        =   39
         Top             =   1400
         Width           =   700
      End
      Begin VB.Label Label87 
         BackColor       =   &H00C0C0C0&
         Caption         =   "               dX, dY"
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
         Left            =   200
         TabIndex        =   38
         Top             =   400
         Width           =   2400
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Õ, Ó"
      Height          =   3500
      Index           =   1
      Left            =   9480
      TabIndex        =   24
      Top             =   0
      Width           =   4500
      Begin VB.ComboBox pnbatXY 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":0080
         Left            =   1100
         List            =   "Pristrelka.frx":008D
         TabIndex        =   28
         Text            =   "1"
         Top             =   900
         Width           =   1000
      End
      Begin VB.CommandButton PrXY 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ðåøèòü"
         Height          =   900
         Index           =   1
         Left            =   2900
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1400
         Width           =   1100
      End
      Begin VB.TextBox pYr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1100
         TabIndex        =   26
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.TextBox pXr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
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
         Top             =   1400
         Width           =   1500
      End
      Begin VB.Label Label94 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ Áàò="
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
         TabIndex        =   32
         Top             =   900
         Width           =   1000
      End
      Begin VB.Label Label92 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Óð="
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
         Left            =   200
         TabIndex        =   31
         Top             =   1900
         Width           =   600
      End
      Begin VB.Label Label91 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Õð="
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
         Left            =   200
         TabIndex        =   30
         Top             =   1400
         Width           =   600
      End
      Begin VB.Label Label87 
         BackColor       =   &H00C0C0C0&
         Caption         =   "                Õ, Ó"
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
         Left            =   200
         TabIndex        =   29
         Top             =   400
         Width           =   2400
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ÍÇÐ"
      Height          =   3500
      Index           =   3
      Left            =   4800
      TabIndex        =   13
      Top             =   0
      Width           =   4500
      Begin VB.ComboBox pnbatNZR 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":009A
         Left            =   2900
         List            =   "Pristrelka.frx":00A7
         TabIndex        =   18
         Text            =   "1"
         Top             =   900
         Width           =   1000
      End
      Begin VB.ComboBox pnkpNZR 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":00B4
         Left            =   1000
         List            =   "Pristrelka.frx":00C7
         TabIndex        =   17
         Text            =   "1"
         Top             =   900
         Width           =   800
      End
      Begin VB.CommandButton PrNZR 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ðåøèòü"
         Height          =   900
         Index           =   3
         Left            =   2800
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1400
         Width           =   1100
      End
      Begin VB.TextBox pdDr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1000
         TabIndex        =   15
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.TextBox pdAr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1000
         TabIndex        =   14
         Text            =   "0"
         Top             =   1400
         Width           =   1500
      End
      Begin VB.Label Label94 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ Áàò="
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
         Left            =   1900
         TabIndex        =   23
         Top             =   900
         Width           =   1000
      End
      Begin VB.Label Label92 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dÄð="
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
         Index           =   3
         Left            =   200
         TabIndex        =   22
         Top             =   1900
         Width           =   700
      End
      Begin VB.Label Label91 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dÀð="
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
         Index           =   3
         Left            =   200
         TabIndex        =   21
         Top             =   1400
         Width           =   700
      End
      Begin VB.Label Label90 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ ÊÏ="
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
         Index           =   3
         Left            =   100
         TabIndex        =   20
         Top             =   900
         Width           =   1000
      End
      Begin VB.Label Label87 
         BackColor       =   &H00C0C0C0&
         Caption         =   "                ÍÇÐ"
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
         Index           =   3
         Left            =   200
         TabIndex        =   19
         Top             =   400
         Width           =   2400
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ÄÀÊ"
      Height          =   3500
      Index           =   0
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   4500
      Begin VB.TextBox pAr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1000
         TabIndex        =   6
         Text            =   "0"
         Top             =   1400
         Width           =   1500
      End
      Begin VB.TextBox pDr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1000
         TabIndex        =   5
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.TextBox pMr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1000
         TabIndex        =   4
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.CommandButton PrDAK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ðåøèòü"
         Height          =   900
         Left            =   2800
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1400
         Width           =   1100
      End
      Begin VB.ComboBox pnkpDAK 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":00DA
         Left            =   1000
         List            =   "Pristrelka.frx":00ED
         TabIndex        =   2
         Text            =   "1"
         Top             =   900
         Width           =   800
      End
      Begin VB.ComboBox pnbatDAK 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Pristrelka.frx":0100
         Left            =   2900
         List            =   "Pristrelka.frx":010D
         TabIndex        =   1
         Text            =   "1"
         Top             =   900
         Width           =   1000
      End
      Begin VB.Label Label87 
         BackColor       =   &H00C0C0C0&
         Caption         =   "                ÄÀÊ"
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
         Index           =   0
         Left            =   200
         TabIndex        =   12
         Top             =   400
         Width           =   2400
      End
      Begin VB.Label Label90 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ ÊÏ="
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
         Index           =   0
         Left            =   100
         TabIndex        =   11
         Top             =   900
         Width           =   1000
      End
      Begin VB.Label Label91 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Àð="
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
         Index           =   0
         Left            =   200
         TabIndex        =   10
         Top             =   1400
         Width           =   600
      End
      Begin VB.Label Label92 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Äð="
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
         Index           =   0
         Left            =   200
         TabIndex        =   9
         Top             =   1900
         Width           =   600
      End
      Begin VB.Label Label93 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ìöð="
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
         Index           =   0
         Left            =   200
         TabIndex        =   8
         Top             =   2400
         Width           =   700
      End
      Begin VB.Label Label94 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¹ Áàò="
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
         Index           =   0
         Left            =   1900
         TabIndex        =   7
         Top             =   900
         Width           =   1000
      End
   End
End
Attribute VB_Name = "Pristrelka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim nb As Single, Alevr As Single, Mclr As Single, nnpl As Single, Apravr As Single, Mcpr As Single, nnpp As Single
 Dim Xl As Single, Yl As Single, hl As Single, Xp As Single, Yp As Single, hp As Single
 nb = pnbatSopr: Alevr = pArLev: Mclr = pMrLev: nnpl = pnkpLev
Apravr = pArPrav: Mcpr = pMrPrav: nnpp = pnkpPrav
If Alevr = 0 And Apravr = 0 Then GoTo 10
   ' SOPR KORREKT
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

  If Alevr < 1500 And Apravr > 4500 Then fir = Abs(Alevr + 6000 - Apravr)
  If Alevr > 4500 And Apravr < 1500 Then fir = Abs(Alevr - (Apravr + 6000))
  If Alevr > Apravr Then fir = Abs(Alevr - Apravr)

  If Alevr < 1500 And Ygolbaz > 4500 Then
        blevr = Abs(Alevr + 6000 - Ygolbaz)
      ElseIf Alevr > 4500 And Ygolbaz < 1500 Then
        blevr = Abs(Alevr - (Ygolbaz + 6000))
      Else
        blevr = Abs(Alevr - Ygolbaz)
  End If

  If Ygolbaz - 3000 < 0 Then
  ybazp = Ygolbaz + 3000
  Else
  ybazp = Ygolbaz - 3000
  End If

  If Apravr < 1500 And ybazp > 4500 Then
  bpravr = Abs(Apravr + 6000 - ybazp)
  ElseIf Apravr > 4500 And ybazp < 1500 Then
  bpravr = Abs(Apravr - (ybazp + 6000))
  Else
  bpravr = Abs(Apravr - ybazp)
  End If

  Dlevr = Abs(baz / (Sin(fir / 100 * 6 * 3.141592 / 180) + 0.001) * Sin(bpravr / 100 * 6 * 3.141592 / 180))
  Dpravr = Abs(baz / (Sin(fir / 100 * 6 * 3.141592 / 180) + 0.001) * Sin(blevr / 100 * 6 * 3.141592 / 180))
  Xcsor = Cos(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Xl
  Ycsor = Sin(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Yl
  If Mclr = 0 Then hr = Mcpr * (Dpravr * 0.001) * 1.05 + hp
  If Mcpr = 0 Then hr = Mclr * (Dlevr * 0.001) * 1.05 + hl
  Xr = Xcsor: Yr = Ycsor
 podRASCHETXY nb, Xr, Yr, hr, dD, dDov, dPr, dYrr, dN
pvKorD.Text = dD: pvKorDov.Text = dDov: pvKorYr.Text = dYrr: pvKorPr.Text = dPr: pvKorTr.Text = dN
10:
End Sub

Private Sub Command2_Click()
 Pristrelka.Hide
End Sub

Private Sub Command3_Click()
RaschSredn.Show
End Sub
Private Sub pAr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pDr.Text = ""
pDr.SetFocus
End If
End Sub
Private Sub pDr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMr.Text = ""
pMr.SetFocus
End If
End Sub
Private Sub pdAr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdDr.Text = ""
pdDr.SetFocus
End If
End Sub
Private Sub pXr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr.Text = ""
pYr.SetFocus
End If
End Sub
Private Sub pdXr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdYr.Text = ""
pdYr.SetFocus
End If
End Sub
Private Sub pArLev_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMrLev.Text = ""
pMrLev.SetFocus
End If
End Sub
Private Sub pArPrav_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMrPrav.Text = ""
pMrPrav.SetFocus
End If
End Sub

  Private Sub PrDAK_Click()
    Dim Arr As Single, Dr As Single, Mr As Single, nkp As Single, nb As Single, Xkp As Single, Ykp As Single, hkp As Single, Xc As Single, Yc As Single, hc As Single
 Dr = pDr: Mr = pMr: nkp = pnkpDAK: nb = pnbatDAK: Arr = pAr
   If nkp = 1 Then Xkp = BP.pXkp1: Ykp = BP.pYkp1: hkp = BP.phkp1
   If nkp = 2 Then Xkp = BP.pXkp2: Ykp = BP.pYkp2: hkp = BP.phkp2
   If nkp = 3 Then Xkp = BP.pXkp3: Ykp = BP.pYkp3: hkp = BP.phkp3
   If nkp = 4 Then Xkp = BP.pXkp4: Ykp = BP.pYkp4: hkp = BP.phkp4
   If nkp = 5 Then Xkp = BP.pXkp5: Ykp = BP.pYkp5: hkp = BP.phkp5
 podRASCHETPRIST Arr, Dr, Xkp, Ykp, Mr, hkp, nb, dD, dDov, dPr, dYrr, dN
 If pvsem = True Then
        If nb = 1 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc: Dt = OZ.pvDt1: Ygolt = OZ.pvYgt1
        If nb = 2 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc: Dt = OZ.pvDt2: Ygolt = OZ.pvYgt2
        If nb = 3 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc: Dt = OZ.pvDt3: Ygolt = OZ.pvYgt3
        ElseIf pkagdomy = True Then
                   If nb = 1 Then Xc = OZzelkagdform.pXc1: Yc = OZzelkagdform.pYc1: hc = OZzelkagdform.phc1: Dt = OZ.pvDt1: Ygolt = OZ.pvYgt1
                    If nb = 2 Then Xc = OZzelkagdform.pXc2: Yc = OZzelkagdform.pYc2: hc = OZzelkagdform.phc2: Dt = OZ.pvDt2: Ygolt = OZ.pvYgt2
                    If nb = 3 Then Xc = OZzelkagdform.pXc3: Yc = OZzelkagdform.pYc3: hc = OZzelkagdform.phc3: Dt = OZ.pvDt3: Ygolt = OZ.pvYgt3
                Else
                        If nb = 1 Then Xc = NZO.pvXce1: Yc = NZO.pvYce1: hc = NZO.pvhc1: Dt = OZ.pvDt1: Ygolt = OZ.pvYgt1
                         If nb = 2 Then Xc = NZO.pvXce2: Yc = NZO.pvYce2: hc = NZO.pvhc2: Dt = OZ.pvDt2: Ygolt = OZ.pvYgt2
                         If nb = 3 Then Xc = NZO.pvXce3: Yc = NZO.pvYce3: hc = NZO.pvhc3: Dt = OZ.pvDt3: Ygolt = OZ.pvYgt3
    End If
   dxc = Xc - Xkp
   dyc = Yc - Ykp
   Dc = Sqr(dxc ^ 2 + dyc ^ 2)
   If Dc = 0 Then GoTo 10
 Ac = Abs(Atn(dyc / (dxc + 0.1)) / 3.141592 * 30) * 100
 If dxc > 0 And dyc > 0 Then Ygolc = Int(Ac)
 If dxc < 0 And dyc > 0 Then Ygolc = Int(3000 - Ac)
 If dxc < 0 And dyc < 0 Then Ygolc = Int(3000 + Ac)
 If dxc > 0 And dyc < 0 Then Ygolc = Int(6000 - Ac)
 If Ygolc < 1500 And Ygolt > 4500 Then
        PS = Ygolc + 6000 - Ygolt
    ElseIf Ygolc > 4500 And Ygolt < 1500 Then
        PS = Ygolc - (Ygolt + 6000)
    Else
        PS = Ygolc - Ygolt
    End If
 Ky = (Round(((Dc + 0.001) / 100) / ((Dt + 0.001) / 1000))) / 10
 SHy = Round(Abs(((PS + 0.001) / 10) / ((Dt + 0.001) / 1000)))
  If PS < 0 Then pvPolOP.Text = "ÑËÅÂÀ"
  If PS > 0 Then pvPolOP.Text = "ÑÏÐÀÂÀ"
pvPS.Text = Round(Abs(PS)): pvSHy.Text = Round(SHy): pvKy.Text = Round(Ky, 1)
pvKorD.Text = dD: pvKorDov.Text = dDov: pvKorYr.Text = dYrr: pvKorPr.Text = dPr: pvKorTr.Text = dN
10:
End Sub

Private Sub PrdXdY_Click(Index As Integer)
Dim nb As Single, dxrc As Single, dyrc As Single, Xc As Single, Yc As Single, Xr As Single, Yr As Single
nb = pnbatdXdY: dxrc = pdXr: dyrc = pdYr
 If pvsem = True Then
        If nb = 1 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc
        If nb = 2 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc
        If nb = 3 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc
        ElseIf pkagdomy = True Then
                   If nb = 1 Then Xc = OZzelkagdform.pXc1: Yc = OZzelkagdform.pYc1: hc = OZzelkagdform.phc1
                    If nb = 2 Then Xc = OZzelkagdform.pXc2: Yc = OZzelkagdform.pYc2: hc = OZzelkagdform.phc2
                    If nb = 3 Then Xc = OZzelkagdform.pXc3: Yc = OZzelkagdform.pYc3: hc = OZzelkagdform.phc3
                Else
        If nb = 1 Then Xc = NZO.pvXce1: Yc = NZO.pvYce1: hc = NZO.pvhc1
        If nb = 2 Then Xc = NZO.pvXce2: Yc = NZO.pvYce2: hc = NZO.pvhc2
        If nb = 3 Then Xc = NZO.pvXce3: Yc = NZO.pvYce3: hc = NZO.pvhc3
    End If
   Xr = Xc + dxrc: Yr = Yc + dyrc
 podRASCHETXY nb, Xr, Yr, hr, dD, dDov, dPr, dYrr, dN
pvKorD.Text = dD: pvKorDov.Text = dDov: pvKorPr.Text = dPr: pvKorTr.Text = dN
End Sub

Private Sub PrNZR_Click(Index As Integer)
Dim nb As Single, nkp As Single, dAr As Single, dDr As Single, Xkp As Single, Ykp As Single, hkp As Single, Xc As Single, Yc As Single, hc As Single, Dt As Single, Ygolt As Single
nb = pnbatNZR: nkp = pnkpNZR
dAr = pdAr: dDr = pdDr
   If nkp = 1 Then Xkp = BP.pXkp1: Ykp = BP.pYkp1: hkp = BP.phkp1
   If nkp = 2 Then Xkp = BP.pXkp2: Ykp = BP.pYkp2: hkp = BP.phkp2
   If nkp = 3 Then Xkp = BP.pXkp3: Ykp = BP.pYkp3: hkp = BP.phkp3
   If nkp = 4 Then Xkp = BP.pXkp4: Ykp = BP.pYkp4: hkp = BP.phkp4
   If nkp = 5 Then Xkp = BP.pXkp5: Ykp = BP.pYkp5: hkp = BP.phkp5
If Xkp = 0 Then GoTo 10
 If pvsem = True Then
        If nb = 1 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc: Dt = OZ.pvDt1: Ygolt = OZ.pvYgt1
        If nb = 2 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc: Dt = OZ.pvDt2: Ygolt = OZ.pvYgt2
        If nb = 3 Then Xc = OZ.pXc: Yc = OZ.pYc: hc = OZ.phc: Dt = OZ.pvDt3: Ygolt = OZ.pvYgt3
        ElseIf pkagdomy = True Then
                   If nb = 1 Then Xc = OZzelkagdform.pXc1: Yc = OZzelkagdform.pYc1: hc = OZzelkagdform.phc1: Dt = OZ.pvDt1: Ygolt = OZ.pvYgt1
                    If nb = 2 Then Xc = OZzelkagdform.pXc2: Yc = OZzelkagdform.pYc2: hc = OZzelkagdform.phc2: Dt = OZ.pvDt2: Ygolt = OZ.pvYgt2
                    If nb = 3 Then Xc = OZzelkagdform.pXc3: Yc = OZzelkagdform.pYc3: hc = OZzelkagdform.phc3: Dt = OZ.pvDt3: Ygolt = OZ.pvYgt3
                Else
                        If nb = 1 Then Xc = NZO.pvXce1: Yc = NZO.pvYce1: hc = NZO.pvhc1: Dt = OZ.pvDt1: Ygolt = OZ.pvYgt1
                         If nb = 2 Then Xc = NZO.pvXce2: Yc = NZO.pvYce2: hc = NZO.pvhc2: Dt = OZ.pvDt2: Ygolt = OZ.pvYgt2
                         If nb = 3 Then Xc = NZO.pvXce3: Yc = NZO.pvYce3: hc = NZO.pvhc3: Dt = OZ.pvDt3: Ygolt = OZ.pvYgt3
    End If
   dxc = Xc - Xkp
   dyc = Yc - Ykp
   Dc = Sqr(dxc ^ 2 + dyc ^ 2)
 Ac = Abs(Atn(dyc / (dxc + 0.1)) / 3.141592 * 30) * 100
 If dxc > 0 And dyc > 0 Then Ygolc = Int(Ac)
 If dxc < 0 And dyc > 0 Then Ygolc = Int(3000 - Ac)
 If dxc < 0 And dyc < 0 Then Ygolc = Int(3000 + Ac)
 If dxc > 0 And dyc < 0 Then Ygolc = Int(6000 - Ac)
 If Ygolc < 1500 And Ygolt > 4500 Then
        PS = Ygolc + 6000 - Ygolt
    ElseIf Ygolc > 4500 And Ygolt < 1500 Then
        PS = Ygolc - (Ygolt + 6000)
    Else
        PS = Ygolc - Ygolt
    End If
 Ky = (Round(((Dc + 0.001) / 100) / ((Dt + 0.001) / 1000))) / 10
 SHy = Round(Abs(((PS + 0.001) / 10) / ((Dt + 0.001) / 1000)))
 Arr = Ygolc + dAr
 Dr = Dc + dDr
  podRASCHETPRIST Arr, Dr, Xkp, Ykp, Mr, hkp, nb, dD, dDov, dPr, dYrr, dN
If PS < 0 Then pvPolOP.Text = "ÑËÅÂÀ"
  If PS > 0 Then pvPolOP.Text = "ÑÏÐÀÂÀ"
pvPS.Text = Round(Abs(PS)): pvSHy.Text = Round(SHy): pvKy.Text = Round(Ky, 1)
pvKorD.Text = dD: pvKorDov.Text = dDov: pvKorPr.Text = dPr: pvKorTr.Text = dN
10:
End Sub
Private Sub PrXY_Click(Index As Integer)
Dim nb As Single, Xr As Single, Yr As Single
        nb = pnbatXY: Xr = pXr: Yr = pYr
podRASCHETXY nb, Xr, Yr, hr, dD, dDov, dPr, dYrr, dN
pvKorD.Text = dD: pvKorDov.Text = dDov: pvKorPr.Text = dPr: pvKorTr.Text = dN
End Sub

Function podRASCHETPRIST(ByVal Arr As Single, ByVal Dr As Single, ByVal Xkp As Single, ByVal Ykp As Single, ByVal Mr As Single, ByVal hkp As Single, ByVal nb As Single, dD, dDov, dPr, dYrr, dN) As Single
Dim vzriv As String
Dim dNtus As Single, Ygoltr As Single, Ygolt As Single, Dt As Single, dXtus As Single
Dim Xb As Single, Yb As Single, hb As Single, hc As Single
152 Xr = Cos(Arr / 100 * 6 * 3.141592 / 180) * Dr + Xkp
 Yr = Sin(Arr / 100 * 6 * 3.141592 / 180) * Dr + Ykp
 hr = (Mr * (Dr * 0.001)) * 1.05 + hkp
 hc = OZ.phc
153
    If nb = 1 Then Xb = BP.pX1: Yb = BP.pY1: hb = BP.ph1
    If Xb = 0 Then
        dXtus = 0: dNtus = 0
            Else
                dXtus = OZ.pvdXtus1: dNtus = OZ.pvdNtus1
    End If
    If nb = 2 Then Xb = BP.pX2: Yb = BP.pY2: hb = BP.ph2: dXtus = OZ.pvdXtus2: dNtus = OZ.pvdNtus2
    If nb = 3 Then Xb = BP.pX3: Yb = BP.pY3: hb = BP.ph3: dXtus = OZ.pvdXtus3: dNtus = OZ.pvdNtus3
        dxr = Xr - Xb
        dyr = Yr - Yb
 Dtr = Sqr(dxr ^ 2 + dyr ^ 2)
 Ar = Abs(Atn(dyr / (dxr + 0.1)) / 3.141592 * 30) * 100
 If dxr > 0 And dyr > 0 Then Ygoltr = Int(Ar)
 If dxr < 0 And dyr > 0 Then Ygoltr = Int(3000 - Ar)
 If dxr < 0 And dyr < 0 Then Ygoltr = Int(3000 + Ar)
 If dxr > 0 And dyr < 0 Then Ygoltr = Int(6000 - Ar)
        If nb = 1 Then Dt = OZ.pvDt1
            If Dt = 0 Then
                Ygolt = 0: vzriv = OZ.pvvzr1
                    Else
                        Ygolt = OZ.pvYgt1: vzriv = OZ.pvvzr1
            End If
        If nb = 2 Then Dt = OZ.pvDt2: Ygolt = OZ.pvYgt2: dXtus = OZ.pvdXtus2: vzriv = OZ.pvvzr2
        If nb = 3 Then Dt = OZ.pvDt3: Ygolt = OZ.pvYgt3: dXtus = OZ.pvdXtus3: vzriv = OZ.pvvzr3
 dD = Round(Dt - Dtr)
 dDov = Round(Ygolt - Ygoltr)
 dPr = Round(dD / (dXtus + 0.01))
        dYrr = Round(((hc - hr + 0.001) / ((Dt + 0.001) * 0.001)) * 0.95)
        dN = Round(dPr * (dNtus + 0.001))
End Function
Function podRASCHETXY(ByVal nb As Single, ByVal Xr As Single, ByVal Yr As Single, ByVal hr As Single, dD, dDov, dPr, dYrr, dN) As Single
'153
Dim Xb As Single, Yb As Single, hb As Single, dXtus As Single, dNtus As Single
Dim Dt As Single, Ygolt As Single
hc = OZ.phc
    If nb = 1 Then Xb = BP.pX1
    If nb = 2 Then Xb = BP.pX2
    If nb = 3 Then Xb = BP.pX3
If Xb = 0 Then GoTo 10
    If nb = 1 Then Xb = BP.pX1: Yb = BP.pY1: hb = BP.ph1: dXtus = OZ.pvdXtus1: dNtus = OZ.pvdNtus1
    If nb = 2 Then Xb = BP.pX2: Yb = BP.pY2: hb = BP.ph2: dXtus = OZ.pvdXtus2: dNtus = OZ.pvdNtus2
    If nb = 3 Then Xb = BP.pX3: Yb = BP.pY3: hb = BP.ph3: dXtus = OZ.pvdXtus3: dNtus = OZ.pvdNtus3
    'If dXtus = 0 Then dXtus = 1
    'If dNtus = 0 Then dNtus = 1
        dxr = Xr - Xb
        dyr = Yr - Yb
 Dtr = Sqr(dxr ^ 2 + dyr ^ 2)
 Ar = Abs(Atn(dyr / (dxr + 0.1)) / 3.141592 * 30) * 100
 If dxr > 0 And dyr > 0 Then Ygoltr = Int(Ar)
 If dxr < 0 And dyr > 0 Then Ygoltr = Int(3000 - Ar)
 If dxr < 0 And dyr < 0 Then Ygoltr = Int(3000 + Ar)
 If dxr > 0 And dyr < 0 Then Ygoltr = Int(6000 - Ar)
        If nb = 1 Then Dt = OZ.pvDt1: Ygolt = OZ.pvYgt1: dXtus = OZ.pvdXtus1: vzriv = OZ.pvvzr1
        If nb = 2 Then Dt = OZ.pvDt2: Ygolt = OZ.pvYgt2: dXtus = OZ.pvdXtus2: vzriv = OZ.pvvzr2
        If nb = 3 Then Dt = OZ.pvDt3: Ygolt = OZ.pvYgt3: dXtus = OZ.pvdXtus3: vzriv = OZ.pvvzr3
 dD = Round(Dt - Dtr)
 dDov = Round(Ygolt - Ygoltr)
 dPr = Round(dD / (dXtus + 0.001))
        dYrr = Round(((hc - hr + 0.001) / ((Dt + 0.001) * 0.001)) * 0.95)
        dN = Round(dPr * (dNtus + 0.001))
10:
End Function
