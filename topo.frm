VERSION 5.00
Begin VB.Form topo 
   BackColor       =   &H0000C0C0&
   Caption         =   "Topo"
   ClientHeight    =   4950
   ClientLeft      =   5310
   ClientTop       =   4155
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
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
   Begin VB.CommandButton SOPR 
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
      Height          =   1200
      Left            =   18720
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2050
      Width           =   1300
   End
   Begin VB.TextBox pMzp 
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
      Left            =   16600
      TabIndex        =   62
      Text            =   "0"
      Top             =   6500
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
      Left            =   16600
      TabIndex        =   61
      Text            =   "0"
      Top             =   5500
      Width           =   1700
   End
   Begin VB.TextBox pMzl 
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
      Left            =   13200
      TabIndex        =   60
      Text            =   "0"
      Top             =   6500
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
      Left            =   13200
      TabIndex        =   59
      Text            =   "0"
      Top             =   5500
      Width           =   1700
   End
   Begin VB.TextBox php 
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
      Left            =   16600
      TabIndex        =   55
      Text            =   "0"
      Top             =   4000
      Width           =   1700
   End
   Begin VB.TextBox pYp 
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
      Left            =   16600
      TabIndex        =   54
      Text            =   "0"
      Top             =   3000
      Width           =   1700
   End
   Begin VB.TextBox pXp 
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
      Left            =   16600
      TabIndex        =   53
      Text            =   "0"
      Top             =   2050
      Width           =   1700
   End
   Begin VB.TextBox phl 
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
      Left            =   13200
      TabIndex        =   52
      Text            =   "0"
      Top             =   4000
      Width           =   1700
   End
   Begin VB.TextBox pYl 
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
      Left            =   13200
      TabIndex        =   51
      Text            =   "0"
      Top             =   3000
      Width           =   1700
   End
   Begin VB.TextBox pXl 
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
      Left            =   13200
      TabIndex        =   50
      Text            =   "0"
      Top             =   2050
      Width           =   1700
   End
   Begin VB.CommandButton OGZ 
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
      Height          =   1200
      Left            =   9800
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5500
      Width           =   1300
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Выход"
      Height          =   1215
      Left            =   18600
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton PGZ 
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
      Height          =   1200
      Left            =   9800
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1080
      Width           =   1300
   End
   Begin VB.TextBox pYr 
      BackColor       =   &H00C0FFC0&
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
      Left            =   950
      TabIndex        =   36
      Text            =   "0"
      Top             =   8500
      Width           =   1700
   End
   Begin VB.TextBox pdovOH 
      BackColor       =   &H00C0FFC0&
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
      Left            =   950
      TabIndex        =   35
      Text            =   "0"
      Top             =   7500
      Width           =   1700
   End
   Begin VB.TextBox pYgt 
      BackColor       =   &H00C0FFC0&
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
      Left            =   950
      TabIndex        =   34
      Text            =   "0"
      Top             =   6500
      Width           =   1700
   End
   Begin VB.TextBox pDt 
      BackColor       =   &H00C0FFC0&
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
      Left            =   950
      TabIndex        =   33
      Text            =   "0"
      Top             =   5500
      Width           =   1700
   End
   Begin VB.TextBox phz 
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
      Left            =   7650
      TabIndex        =   27
      Text            =   "0"
      Top             =   7500
      Width           =   1700
   End
   Begin VB.TextBox pYz 
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
      Left            =   7650
      TabIndex        =   26
      Text            =   "0"
      Top             =   6500
      Width           =   1700
   End
   Begin VB.TextBox pXz 
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
      Left            =   7650
      TabIndex        =   25
      Text            =   "0"
      Top             =   5500
      Width           =   1700
   End
   Begin VB.TextBox pMz 
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
      Left            =   7650
      TabIndex        =   21
      Text            =   "0"
      Top             =   3000
      Width           =   1700
   End
   Begin VB.TextBox pD 
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
      Left            =   7650
      TabIndex        =   20
      Text            =   "0"
      Top             =   2050
      Width           =   1700
   End
   Begin VB.TextBox pA 
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
      Left            =   7650
      TabIndex        =   19
      Text            =   "0"
      Top             =   1100
      Width           =   1700
   End
   Begin VB.TextBox phnp 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   4400
      TabIndex        =   14
      Text            =   "0"
      Top             =   3000
      Width           =   1700
   End
   Begin VB.TextBox pYnp 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   4400
      TabIndex        =   13
      Text            =   "0"
      Top             =   2050
      Width           =   1700
   End
   Begin VB.TextBox pXnp 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   4400
      TabIndex        =   12
      Text            =   "0"
      Top             =   1100
      Width           =   1700
   End
   Begin VB.TextBox pOH 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   950
      TabIndex        =   7
      Text            =   "0"
      Top             =   4000
      Width           =   1700
   End
   Begin VB.TextBox phop 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   950
      TabIndex        =   4
      Text            =   "0"
      Top             =   3000
      Width           =   1700
   End
   Begin VB.TextBox pYop 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   950
      TabIndex        =   3
      Text            =   "0"
      Top             =   2050
      Width           =   1700
   End
   Begin VB.TextBox pXop 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   950
      TabIndex        =   2
      Text            =   "0"
      Top             =   1100
      Width           =   1700
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "             ОП"
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   250
      TabIndex        =   66
      Top             =   250
      Width           =   2400
   End
   Begin VB.Label Label36 
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
      Left            =   16000
      TabIndex        =   64
      Top             =   6500
      Width           =   600
   End
   Begin VB.Label Label35 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ап="
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
      Left            =   16000
      TabIndex        =   63
      Top             =   5500
      Width           =   500
   End
   Begin VB.Label Label34 
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
      Left            =   12500
      TabIndex        =   58
      Top             =   6500
      Width           =   600
   End
   Begin VB.Label Label33 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ал="
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
      Left            =   12500
      TabIndex        =   57
      Top             =   5500
      Width           =   500
   End
   Begin VB.Label Label32 
      BackColor       =   &H0000C0C0&
      Caption         =   "             Засечка"
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
      Left            =   13800
      TabIndex        =   56
      Top             =   4800
      Width           =   2500
   End
   Begin VB.Label Label31 
      BackColor       =   &H0000C0C0&
      Caption         =   "h="
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
      Left            =   16000
      TabIndex        =   49
      Top             =   4000
      Width           =   400
   End
   Begin VB.Label Label30 
      BackColor       =   &H0000C0C0&
      Caption         =   "У="
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
      Left            =   16000
      TabIndex        =   48
      Top             =   3120
      Width           =   400
   End
   Begin VB.Label Label29 
      BackColor       =   &H0000C0C0&
      Caption         =   "Х="
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
      Left            =   16000
      TabIndex        =   47
      Top             =   2050
      Width           =   400
   End
   Begin VB.Label Label28 
      BackColor       =   &H0000C0C0&
      Caption         =   "h="
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
      Left            =   12500
      TabIndex        =   46
      Top             =   4000
      Width           =   400
   End
   Begin VB.Label Label27 
      BackColor       =   &H0000C0C0&
      Caption         =   "У="
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
      Left            =   12500
      TabIndex        =   45
      Top             =   3000
      Width           =   400
   End
   Begin VB.Label Label26 
      BackColor       =   &H0000C0C0&
      Caption         =   "Х="
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
      Left            =   12500
      TabIndex        =   44
      Top             =   2050
      Width           =   400
   End
   Begin VB.Label Label25 
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
      Height          =   400
      Left            =   16000
      TabIndex        =   43
      Top             =   1100
      Width           =   2300
   End
   Begin VB.Label Label24 
      BackColor       =   &H0000C0C0&
      Caption         =   "           Левый"
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
      Left            =   12500
      TabIndex        =   42
      Top             =   1100
      Width           =   2300
   End
   Begin VB.Label Label23 
      BackColor       =   &H0000C0C0&
      Caption         =   "              Сопряженка"
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
      Left            =   13800
      TabIndex        =   41
      Top             =   360
      Width           =   3400
   End
   Begin VB.Label Label22 
      BackColor       =   &H0000C0C0&
      Caption         =   "   Прямоугольные"
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
      Left            =   7080
      TabIndex        =   40
      Top             =   4800
      Width           =   2500
   End
   Begin VB.Label Label21 
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
      Left            =   250
      TabIndex        =   32
      Top             =   8500
      Width           =   600
   End
   Begin VB.Label Label20 
      BackColor       =   &H0000C0C0&
      Caption         =   "ОН"
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
      Left            =   250
      TabIndex        =   31
      Top             =   7500
      Width           =   500
   End
   Begin VB.Label Label19 
      BackColor       =   &H0000C0C0&
      Caption         =   "Угт="
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
      Left            =   200
      TabIndex        =   30
      Top             =   6500
      Width           =   700
   End
   Begin VB.Label Label18 
      BackColor       =   &H0000C0C0&
      Caption         =   "Дт="
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
      Left            =   250
      TabIndex        =   29
      Top             =   5500
      Width           =   600
   End
   Begin VB.Label Label17 
      BackColor       =   &H0000C0C0&
      Caption         =   " Топо данные по цели"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   28
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label Label16 
      BackColor       =   &H0000C0C0&
      Caption         =   "hц="
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
      Left            =   7000
      TabIndex        =   24
      Top             =   7500
      Width           =   500
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000C0C0&
      Caption         =   "Уц="
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
      Left            =   7000
      TabIndex        =   23
      Top             =   6500
      Width           =   500
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      Caption         =   "Хц="
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
      Left            =   7000
      TabIndex        =   22
      Top             =   5500
      Width           =   500
   End
   Begin VB.Label Label13 
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
      Left            =   7000
      TabIndex        =   18
      Top             =   3000
      Width           =   600
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C0C0&
      Caption         =   "Д="
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
      Left            =   7000
      TabIndex        =   17
      Top             =   2050
      Width           =   400
   End
   Begin VB.Label Label11 
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
      Left            =   7000
      TabIndex        =   16
      Top             =   1100
      Width           =   400
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "          Полярные"
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
      Left            =   7000
      TabIndex        =   15
      Top             =   250
      Width           =   2400
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "h="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   3700
      TabIndex        =   11
      Top             =   3000
      Width           =   400
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "У="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   3700
      TabIndex        =   10
      Top             =   2050
      Width           =   400
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "Х="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   3700
      TabIndex        =   9
      Top             =   1100
      Width           =   400
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "                  НП"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   3720
      TabIndex        =   8
      Top             =   250
      Width           =   2400
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "OH="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   250
      TabIndex        =   6
      Top             =   4000
      Width           =   600
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "h="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   250
      TabIndex        =   5
      Top             =   3000
      Width           =   400
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "У="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   250
      TabIndex        =   1
      Top             =   2050
      Width           =   400
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Х="
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   500
      Left            =   250
      TabIndex        =   0
      Top             =   1100
      Width           =   400
   End
End
Attribute VB_Name = "topo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public A As Single
Public D As Single
Public Mz As Single
Public Xnp As Single
Public Ynp As Single
Public hnp As Single
Public Xz As Single
Public Yz As Single
Public hz As Single
Public Xop As Single
Public Yop As Single
Public hop As Single
Dim dx As Single
Dim dy As Single
Public Dt As Single
Public Yr As Single
Public dh, Pi, Ygolt, OH, dovort As Single
Dim ar As Single
Dim dxso, Xp, Xl, dyso, Yp, Yl, baz, aso, Ygolbaz, Alev, Aprav, fi, blev, ybazp, bprav, Dlev, Dprav, Mcl, Mcp, hp, hl As Single

Private Sub Command2_Click()
End
End Sub

Private Sub OGZ_Click()
Ygolt = 0
Xop = pXop: Yop = pYop: hop = phop: OH = pOH
Xz = pXz: Yz = pYz: hz = phz
Pi = 3.14159265358
dx = Xz - Xop: dy = Yz - Yop: dh = hz - hop
Dt = Sqr(dx ^ 2 + dy ^ 2)
Yr = (dh / (Dt * 0.001 + 0.001)) * 0.95
 ar = Abs(Atn(dy / (dx + 0.001)) / Pi * 30) * 100
If dx > 0 And dy > 0 Then Ygolt = ar
If dx < 0 And dy > 0 Then Ygolt = 3000 - ar
If dx < 0 And dy < 0 Then Ygolt = 3000 + ar
If dx > 0 And dy < 0 Then Ygolt = 6000 - ar
If Ygolt <= 1500 And OH >= 4500 Then
      dovort = Ygolt + 6000 - OH
      ElseIf OH <= 1500 And Ygolt >= 4500 Then
      dovort = Ygolt - (OH + 6000)
      Else
      dovort = Ygolt - OH
      End If
pDt.Text = Format(Dt, "0")
pYgt.Text = Format(Ygolt, "0")
pdovOH.Text = Format(dovort, "0")
pYr.Text = Format(Yr, "0")
End Sub

Private Sub PGZ_Click()
Ygolt = 0
Xop = pXop: Yop = pYop: hop = phop: OH = pOH
A = pA: D = pD: Mz = pMz: Xnp = pXnp: Ynp = pYnp: hnp = phnp
Xz = Cos((A + 0.001) / 100 * 6 * 3.141592 / 180) * D + Xnp
Yz = Sin((A + 0.001) / 100 * 6 * 3.141592 / 180) * D + Ynp
hz = Mz * ((D + 0.001) / 1000) * 1.05 + hnp
pXz = Format(Xz, "0")
pYz = Format(Yz, "0")
phz = Format(hz, "0")
Pi = 3.14159265358
dx = Xz - Xop: dy = Yz - Yop: dh = hz - hop
Dt = Sqr(dx ^ 2 + dy ^ 2)
Yr = (dh / (Dt * 0.001 + 0.001)) * 0.95
 ar = Abs(Atn(dy / (dx + 0.001)) / Pi * 30) * 100
If dx > 0 And dy > 0 Then Ygolt = ar
If dx < 0 And dy > 0 Then Ygolt = 3000 - ar
If dx < 0 And dy < 0 Then Ygolt = 3000 + ar
If dx > 0 And dy < 0 Then Ygolt = 6000 - ar
If Ygolt <= 1500 And OH >= 4500 Then
      dovort = Ygolt + 6000 - OH
      ElseIf OH <= 1500 And Ygolt >= 4500 Then
      dovort = Ygolt - (OH + 6000)
      Else
      dovort = Ygolt - OH
      End If
pDt.Text = Format(Dt, "0")
pYgt.Text = Format(Ygolt, "0")
pdovOH.Text = Format(dovort, "0")
pYr.Text = Format(Yr, "0")
End Sub



Private Sub SOPR_Click()
Ygolt = 0
Xop = pXop: Yop = pYop: hop = phop: OH = pOH
Xp = pXp: Yp = pYp: hp = php
Xl = pXl: Yl = pYl: hl = phl
Alev = pAl: Aprav = pAp
Mcp = pMzp: Mcl = pMzl
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
  Xz = Cos(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Xl
  Yz = Sin(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Yl
  If Mcl = 0 Then hz = Mcp * (Dprav * 0.001) * 1.05 + hp
  If Mcp = 0 Then hz = Mcl * (Dlev * 0.001) * 1.05 + hl
  pXz = Format(Xz, "0")
pYz = Format(Yz, "0")
phz = Format(hz, "0")
Pi = 3.14159265358
dx = Xz - Xop: dy = Yz - Yop: dh = hz - hop
Dt = Sqr(dx ^ 2 + dy ^ 2)
Yr = (dh / (Dt * 0.001 + 0.001)) * 0.95
 ar = Abs(Atn(dy / (dx + 0.001)) / Pi * 30) * 100
If dx > 0 And dy > 0 Then Ygolt = ar
If dx < 0 And dy > 0 Then Ygolt = 3000 - ar
If dx < 0 And dy < 0 Then Ygolt = 3000 + ar
If dx > 0 And dy < 0 Then Ygolt = 6000 - ar
If Ygolt <= 1500 And OH >= 4500 Then
      dovort = Ygolt + 6000 - OH
      ElseIf OH <= 1500 And Ygolt >= 4500 Then
      dovort = Ygolt - (OH + 6000)
      Else
      dovort = Ygolt - OH
      End If
pDt.Text = Format(Dt, "0")
pYgt.Text = Format(Ygolt, "0")
pdovOH.Text = Format(dovort, "0")
pYr.Text = Format(Yr, "0")
End Sub
