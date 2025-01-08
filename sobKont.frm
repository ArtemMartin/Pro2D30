VERSION 5.00
Begin VB.Form SobKontrol 
   BackColor       =   &H00808080&
   Caption         =   "Бусоль огня"
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   14235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton pol3B 
      BackColor       =   &H000080FF&
      Caption         =   "3-я Бат"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2100
      Width           =   1000
   End
   Begin VB.CommandButton pol2B 
      BackColor       =   &H000080FF&
      Caption         =   "2-я Бат"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CommandButton pol1B 
      BackColor       =   &H000080FF&
      Caption         =   "1-я Бат"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   300
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "ВЫХОД"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   12400
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   7000
      Width           =   1400
   End
   Begin VB.TextBox pOshOr1 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   55
      Text            =   "0"
      Top             =   8300
      Width           =   1000
   End
   Begin VB.TextBox pOshOr2 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9300
      TabIndex        =   54
      Text            =   "0"
      Top             =   8280
      Width           =   1000
   End
   Begin VB.TextBox pOshOr3 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7800
      TabIndex        =   53
      Text            =   "0"
      Top             =   8280
      Width           =   1000
   End
   Begin VB.TextBox pOshOr4 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6300
      TabIndex        =   52
      Text            =   "0"
      Top             =   8280
      Width           =   1000
   End
   Begin VB.TextBox pOshOr5 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4800
      TabIndex        =   51
      Text            =   "0"
      Top             =   8300
      Width           =   1000
   End
   Begin VB.TextBox pOshOr6 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3300
      TabIndex        =   50
      Text            =   "0"
      Top             =   8300
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   2400
      Left            =   12240
      TabIndex        =   45
      Top             =   3800
      Width           =   1455
      Begin VB.OptionButton pOsnovnoe3 
         BackColor       =   &H00808080&
         Caption         =   "3-й"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   1560
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton pOsnovnoe2 
         BackColor       =   &H00808080&
         Caption         =   "2-й"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox pDyKont1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   44
      Text            =   "0"
      Top             =   7200
      Width           =   1000
   End
   Begin VB.TextBox pDyKont2 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9300
      TabIndex        =   43
      Text            =   "0"
      Top             =   7200
      Width           =   1000
   End
   Begin VB.TextBox pDyKont3 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7800
      TabIndex        =   42
      Text            =   "0"
      Top             =   7200
      Width           =   1000
   End
   Begin VB.TextBox pDyKont4 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6300
      TabIndex        =   41
      Text            =   "0"
      Top             =   7200
      Width           =   1000
   End
   Begin VB.TextBox pDyKont5 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4800
      TabIndex        =   40
      Text            =   "0"
      Top             =   7200
      Width           =   1000
   End
   Begin VB.TextBox pDyKont6 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3300
      TabIndex        =   39
      Text            =   "0"
      Top             =   7200
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   7200
      TabIndex        =   33
      Top             =   3800
      Width           =   3855
      Begin VB.TextBox pVeer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2160
         TabIndex        =   36
         Text            =   "0"
         Top             =   550
         Width           =   600
      End
      Begin VB.OptionButton flRazdelit 
         BackColor       =   &H00808080&
         Caption         =   "Разделить"
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton flSoedinit 
         BackColor       =   &H00808080&
         Caption         =   "Соединить"
         Height          =   375
         Left            =   360
         TabIndex        =   34
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox pDovorot 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5000
      TabIndex        =   32
      Text            =   "0"
      Top             =   4100
      Width           =   1000
   End
   Begin VB.TextBox pON 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3000
      TabIndex        =   31
      Text            =   "0"
      Top             =   4100
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "РЕШИТЬ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9300
      Width           =   13455
   End
   Begin VB.TextBox pDyoryd1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   26
      Text            =   "0"
      Top             =   5900
      Width           =   1000
   End
   Begin VB.TextBox pDyoryd2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9300
      TabIndex        =   25
      Text            =   "0"
      Top             =   5900
      Width           =   1000
   End
   Begin VB.TextBox pDyoryd3 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7800
      TabIndex        =   24
      Text            =   "0"
      Top             =   5900
      Width           =   1000
   End
   Begin VB.TextBox pDyoryd4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6300
      TabIndex        =   23
      Text            =   "0"
      Top             =   5900
      Width           =   1000
   End
   Begin VB.TextBox pDyoryd5 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4800
      TabIndex        =   22
      Text            =   "0"
      Top             =   5900
      Width           =   1000
   End
   Begin VB.TextBox pDyoryd6 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3300
      TabIndex        =   21
      Text            =   "0"
      Top             =   5900
      Width           =   1000
   End
   Begin VB.TextBox pYglpobys1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   19
      Text            =   "0"
      Top             =   2000
      Width           =   1000
   End
   Begin VB.TextBox pYglpobys2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9300
      TabIndex        =   18
      Text            =   "0"
      Top             =   2000
      Width           =   1000
   End
   Begin VB.TextBox pYglpobys3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7800
      TabIndex        =   17
      Text            =   "0"
      Top             =   2000
      Width           =   1000
   End
   Begin VB.TextBox pYglpobys4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6300
      TabIndex        =   16
      Text            =   "0"
      Top             =   2000
      Width           =   1000
   End
   Begin VB.TextBox pYglpobys5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4800
      TabIndex        =   15
      Text            =   "0"
      Top             =   2000
      Width           =   1000
   End
   Begin VB.TextBox pYglpobys6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3300
      TabIndex        =   14
      Text            =   "0"
      Top             =   2000
      Width           =   1000
   End
   Begin VB.TextBox pDYor1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   13
      Text            =   "0"
      Top             =   900
      Width           =   1000
   End
   Begin VB.TextBox pDYor2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9300
      TabIndex        =   12
      Text            =   "0"
      Top             =   900
      Width           =   1000
   End
   Begin VB.TextBox pDYor3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7800
      TabIndex        =   11
      Text            =   "0"
      Top             =   900
      Width           =   1000
   End
   Begin VB.TextBox pDYor4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6300
      TabIndex        =   10
      Text            =   "0"
      Top             =   900
      Width           =   1000
   End
   Begin VB.TextBox pDYor5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4800
      TabIndex        =   9
      Text            =   "0"
      Top             =   900
      Width           =   1000
   End
   Begin VB.TextBox pDYor6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3300
      TabIndex        =   8
      Text            =   "0"
      Top             =   900
      Width           =   1000
   End
   Begin VB.Label Label16 
      BackColor       =   &H00808080&
      Caption         =   "Ошибки"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   700
      TabIndex        =   49
      Top             =   8350
      Width           =   1300
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808080&
      Caption         =   "Основное"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12100
      TabIndex        =   46
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808080&
      Caption         =   "Контрольные данные"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   50
      TabIndex        =   38
      Top             =   7100
      Width           =   2000
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808080&
      Caption         =   "Установки по цели"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   50
      TabIndex        =   37
      Top             =   4000
      Width           =   1600
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      Caption         =   "Команда в веер"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8040
      TabIndex        =   30
      Top             =   3000
      Width           =   1600
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      Caption         =   "Доворот на цель"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4800
      TabIndex        =   29
      Top             =   3000
      Width           =   1600
   End
   Begin VB.Label Label10 
      BackColor       =   &H00808080&
      Caption         =   "ОН стрельбы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   28
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      Caption         =   "Диркцыонные углы орудий"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   100
      TabIndex        =   20
      Top             =   5800
      Width           =   2200
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Caption         =   "1-й"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11000
      TabIndex        =   7
      Top             =   50
      Width           =   600
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Caption         =   "2-й"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9500
      TabIndex        =   6
      Top             =   50
      Width           =   600
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      Caption         =   "3-й"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8000
      TabIndex        =   5
      Top             =   50
      Width           =   600
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "4-й"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6500
      TabIndex        =   4
      Top             =   50
      Width           =   600
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "5-й"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5000
      TabIndex        =   3
      Top             =   50
      Width           =   600
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "6-й"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3500
      TabIndex        =   2
      Top             =   50
      Width           =   600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Угломеры по буссоли"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   100
      TabIndex        =   1
      Top             =   2000
      Width           =   2200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Дирекцыонные по панорамам"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   100
      TabIndex        =   0
      Top             =   700
      Width           =   2200
   End
End
Attribute VB_Name = "SobKontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim dyOr1 As Double
    Dim dyOr2 As Double
    Dim dyOr3 As Double
    Dim dyOr4 As Double
    Dim dyOr5 As Double
    Dim dyOr6 As Double
    Dim yglOr1 As Double
    Dim yglOr2 As Double
    Dim yglOr3 As Double
    Dim yglOr4 As Double
    Dim yglOr5 As Double
    Dim yglOr6 As Double
    Dim dyNavOryd1 As Double
    Dim dyNavOryd2 As Double
    Dim dyNavOryd3 As Double
    Dim dyNavOryd4 As Double
    Dim dyNavOryd5 As Double
    Dim dyNavOryd6 As Double
    Dim osnNapr As Integer, dovorot As Integer
    Dim kont1Or As Integer, kont2Or As Integer, kont3Or As Integer, kont4Or As Integer, kont5Or As Integer, kont6Or As Integer
    Dim osnovnoe As Integer, veer As Integer, soedrazd As Integer, konOr As Integer
    Dim osh1 As Integer, osh2 As Integer, osh3 As Integer, osh4 As Integer, osh5 As Integer, osh6 As Integer
    If pOsnovnoe2 = True Then
        osnovnoe = 2
    Else
        osnovnoe = 3
    End If
    
    If flSoedinit = True Then
        soedrazd = 0
    Else
        soedrazd = 1
    End If
            
    dyOr1 = pDYor1
    dyOr2 = pDYor2
    dyOr3 = pDYor3
    dyOr4 = pDYor4
    dyOr5 = pDYor5
    dyOr6 = pDYor6
    yglOr1 = pYglpobys1
    yglOr2 = pYglpobys2
    yglOr3 = pYglpobys3
    yglOr4 = pYglpobys4
    yglOr5 = pYglpobys5
    yglOr6 = pYglpobys6
    
    If dyOr1 + yglOr1 > 6000 Then
        dyNavOryd1 = dyOr1 + yglOr1 - 6000
        Else
        dyNavOryd1 = dyOr1 + yglOr1
    End If
    If dyOr2 + yglOr2 > 6000 Then
        dyNavOryd2 = dyOr2 + yglOr2 - 6000
        Else
        dyNavOryd2 = dyOr2 + yglOr2
    End If
    If dyOr3 + yglOr3 > 6000 Then
        dyNavOryd3 = dyOr3 + yglOr3 - 6000
        Else
        dyNavOryd3 = dyOr3 + yglOr3
    End If
    If dyOr4 + yglOr4 > 6000 Then
        dyNavOryd4 = dyOr4 + yglOr4 - 6000
        Else
        dyNavOryd4 = dyOr4 + yglOr4
    End If
    If dyOr5 + yglOr5 > 6000 Then
        dyNavOryd5 = dyOr5 + yglOr5 - 6000
        Else
        dyNavOryd5 = dyOr5 + yglOr5
    End If
    If dyOr6 + yglOr6 > 6000 Then
        dyNavOryd6 = dyOr6 + yglOr6 - 6000
        Else
        dyNavOryd6 = dyOr6 + yglOr6
    End If
    pDyoryd1.Text = dyNavOryd1
    pDyoryd2.Text = dyNavOryd2
    pDyoryd3.Text = dyNavOryd3
    pDyoryd4.Text = dyNavOryd4
    pDyoryd5.Text = dyNavOryd5
    pDyoryd6.Text = dyNavOryd6
    
    osnNapr = pON: dovorot = pDovorot: veer = pVeer
    kont1Or = osnNapr + pDovorot
    kont2Or = osnNapr + pDovorot
    kont3Or = osnNapr + pDovorot
    kont4Or = osnNapr + pDovorot
    kont5Or = osnNapr + pDovorot
    kont6Or = osnNapr + pDovorot
    
    If dyOr1 <> 0 Then a = polKontDY(1, kont1Or, osnovnoe, veer, soedrazd, konOr)
    kont1Or = konOr
    If dyOr2 <> 0 Then a = polKontDY(2, kont2Or, osnovnoe, veer, soedrazd, konOr)
    kont2Or = konOr
    If dyOr3 <> 0 Then a = polKontDY(3, kont3Or, osnovnoe, veer, soedrazd, konOr)
    kont3Or = konOr
    If dyOr4 <> 0 Then a = polKontDY(4, kont4Or, osnovnoe, veer, soedrazd, konOr)
    kont4Or = konOr
    If dyOr5 <> 0 Then a = polKontDY(5, kont5Or, osnovnoe, veer, soedrazd, konOr)
    kont5Or = konOr
    If dyOr6 <> 0 Then a = polKontDY(6, kont6Or, osnovnoe, veer, soedrazd, konOr)
    kont6Or = konOr

    If dyOr1 <> 0 Then pDyKont1.Text = kont1Or
    If dyOr2 <> 0 Then pDyKont2.Text = kont2Or
    If dyOr3 <> 0 Then pDyKont3.Text = kont3Or
    If dyOr4 <> 0 Then pDyKont4.Text = kont4Or
    If dyOr5 <> 0 Then pDyKont5.Text = kont5Or
    If dyOr6 <> 0 Then pDyKont6.Text = kont6Or
    
    If dyOr1 <> 0 Then osh1 = kont1Or - dyNavOryd1
    If dyOr2 <> 0 Then osh2 = kont2Or - dyNavOryd2
    If dyOr3 <> 0 Then osh3 = kont3Or - dyNavOryd3
    If dyOr4 <> 0 Then osh4 = kont4Or - dyNavOryd4
    If dyOr5 <> 0 Then osh5 = kont5Or - dyNavOryd5
    If dyOr6 <> 0 Then osh6 = kont6Or - dyNavOryd6
    
    If dyOr1 <> 0 Then pOshOr1.Text = osh1
    If dyOr2 <> 0 Then pOshOr2.Text = osh2
    If dyOr3 <> 0 Then pOshOr3.Text = osh3
    If dyOr4 <> 0 Then pOshOr4.Text = osh4
    If dyOr5 <> 0 Then pOshOr5.Text = osh5
    If dyOr6 <> 0 Then pOshOr6.Text = osh6
    
End Sub

Private Sub Command2_Click()
SobKontrol.Hide
End Sub

Private Sub pDYor1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pDYor2.Text = ""
   pDYor2.SetFocus
Else
End If
End Sub
Private Sub pDYor2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pDYor3.Text = ""
   pDYor3.SetFocus
Else
End If
End Sub
Private Sub pDYor3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pDYor4.Text = ""
   pDYor4.SetFocus
Else
End If
End Sub
Private Sub pDYor4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pDYor5.Text = ""
   pDYor5.SetFocus
Else
End If
End Sub
Private Sub pDYor5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pDYor6.Text = ""
   pDYor6.SetFocus
Else
End If
End Sub

Private Sub pol1B_Click()
pON.Text = BP.pOH1
pDovorot.Text = OZ.pvDov1
End Sub

Private Sub pol2B_Click()
pON.Text = BP.pOH2
pDovorot.Text = OZ.pvDov2
End Sub

Private Sub pol3B_Click()
pON.Text = BP.pOH3
pDovorot.Text = OZ.pvDov3
End Sub

Private Sub pON_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pDovorot.Text = ""
    pDovorot.SetFocus
Else
End If
End Sub

Private Sub pYglpobys1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pYglpobys2.Text = ""
    pYglpobys2.SetFocus
Else
End If
End Sub
Private Sub pYglpobys2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pYglpobys3.Text = ""
    pYglpobys3.SetFocus
Else
End If
End Sub
Private Sub pYglpobys3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pYglpobys4.Text = ""
    pYglpobys4.SetFocus
Else
End If
End Sub
Private Sub pYglpobys4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pYglpobys5.Text = ""
    pYglpobys5.SetFocus
Else
End If
End Sub
Private Sub pYglpobys5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pYglpobys6.Text = ""
    pYglpobys6.SetFocus
Else
End If
End Sub
Function polKontDY(ByVal nomerOryd As Integer, ByVal dyOryd As Integer, ByVal osnovnoe As Integer, ByVal veer As Integer, ByVal soedrazd As Integer, konOr) As Integer
If osnovnoe = 2 Then
    If soedrazd = 0 Then
        If nomerOryd = 1 Then
        dyOryd = dyOryd - veer
        ElseIf nomerOryd = 2 Then
        dyOryd = dyOryd
        ElseIf nomerOryd = 3 Then
        dyOryd = dyOryd + veer
        ElseIf nomerOryd = 4 Then
        dyOryd = dyOryd + (veer * 2)
        ElseIf nomerOryd = 5 Then
        dyOryd = dyOryd + (veer * 3)
        Else
        End If
    Else
    If nomerOryd = 1 Then
        dyOryd = dyOryd + veer
        ElseIf nomerOryd = 2 Then
        dyOryd = dyOryd
        ElseIf nomerOryd = 3 Then
        dyOryd = dyOryd - veer
        ElseIf nomerOryd = 4 Then
        dyOryd = dyOryd - (veer * 2)
        ElseIf nomerOryd = 5 Then
        dyOryd = dyOryd - (veer * 3)
        Else
        End If
    End If
Else
If soedrazd = 0 Then
        If nomerOryd = 1 Then
        dyOryd = dyOryd - (veer * 2)
        ElseIf nomerOryd = 2 Then
        dyOryd = dyOryd - veer
        ElseIf nomerOryd = 3 Then
        dyOryd = dyOryd
        ElseIf nomerOryd = 4 Then
        dyOryd = dyOryd + veer
        ElseIf nomerOryd = 5 Then
        dyOryd = dyOryd + (veer * 2)
        ElseIf nomerOryd = 6 Then
        dyOryd = dyOryd + (veer * 3)
        Else
        End If
    Else
    If nomerOryd = 1 Then
        dyOryd = dyOryd + (veer * 2)
        ElseIf nomerOryd = 2 Then
        dyOryd = dyOryd + veer
        ElseIf nomerOryd = 3 Then
        dyOryd = dyOryd
        ElseIf nomerOryd = 4 Then
        dyOryd = dyOryd - veer
        ElseIf nomerOryd = 5 Then
        dyOryd = dyOryd - (veer * 2)
         ElseIf nomerOryd = 6 Then
        dyOryd = dyOryd - (veer * 3)
        Else
        End If
    End If
End If
konOr = dyOryd
End Function
