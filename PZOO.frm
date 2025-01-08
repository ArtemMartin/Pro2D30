VERSION 5.00
Begin VB.Form PZO 
   Caption         =   "ПЗО"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   14115
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Рубежи"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9030
      Left            =   8175
      TabIndex        =   37
      Top             =   100
      Width           =   11895
      Begin VB.TextBox pvFr3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   73
         Text            =   "0"
         Top             =   8300
         Width           =   1000
      End
      Begin VB.TextBox pvtip3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   72
         Text            =   "ФРОНТ"
         Top             =   7700
         Width           =   1000
      End
      Begin VB.TextBox pvFr2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2300
         TabIndex        =   71
         Text            =   "0"
         Top             =   8300
         Width           =   1000
      End
      Begin VB.TextBox pvtip2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2300
         TabIndex        =   70
         Text            =   "ФРОНТ"
         Top             =   7700
         Width           =   1000
      End
      Begin VB.TextBox pvFr1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1000
         TabIndex        =   69
         Text            =   "0"
         Top             =   8300
         Width           =   1000
      End
      Begin VB.TextBox pvtip1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1000
         TabIndex        =   68
         Text            =   "ФРОНТ"
         Top             =   7700
         Width           =   1000
      End
      Begin VB.CommandButton navesti6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Навести"
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
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   6500
         Width           =   1500
      End
      Begin VB.TextBox pYz6 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2900
         TabIndex        =   61
         Text            =   "0"
         Top             =   6000
         Width           =   1000
      End
      Begin VB.TextBox pXz6 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1700
         TabIndex        =   60
         Text            =   "0"
         Top             =   6000
         Width           =   1000
      End
      Begin VB.CommandButton navesti5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Навести"
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
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5400
         Width           =   1500
      End
      Begin VB.TextBox pYz5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2900
         TabIndex        =   57
         Text            =   "0"
         Top             =   4900
         Width           =   1000
      End
      Begin VB.TextBox pXz5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1700
         TabIndex        =   56
         Text            =   "0"
         Top             =   4900
         Width           =   1000
      End
      Begin VB.CommandButton navesti4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Навести"
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
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4300
         Width           =   1500
      End
      Begin VB.TextBox pYz4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2900
         TabIndex        =   53
         Text            =   "0"
         Top             =   3800
         Width           =   1000
      End
      Begin VB.TextBox pXz4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1700
         TabIndex        =   52
         Text            =   "0"
         Top             =   3800
         Width           =   1000
      End
      Begin VB.CommandButton navesti3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Навести"
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
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3200
         Width           =   1500
      End
      Begin VB.TextBox pYz3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2900
         TabIndex        =   49
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.TextBox pXz3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1700
         TabIndex        =   48
         Text            =   "0"
         Top             =   2700
         Width           =   1000
      End
      Begin VB.CommandButton navesti2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Навести"
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
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2100
         Width           =   1500
      End
      Begin VB.TextBox pYz2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2900
         TabIndex        =   45
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pXz2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1700
         TabIndex        =   44
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.CommandButton navesti1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Навести"
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
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pYz1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2900
         TabIndex        =   41
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pXz1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1700
         TabIndex        =   39
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3200
         TabIndex        =   79
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Х"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2000
         TabIndex        =   78
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   3 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   67
         Top             =   7200
         Width           =   1000
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   2 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2300
         TabIndex        =   66
         Top             =   7200
         Width           =   1000
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C0C0&
         Caption         =   "   1 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1000
         TabIndex        =   65
         Top             =   7200
         Width           =   1000
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Бат уч"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   100
         TabIndex        =   64
         Top             =   8300
         Width           =   800
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
         Caption         =   "тип ЗО"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   100
         TabIndex        =   63
         Top             =   7700
         Width           =   800
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Рубеж №6 Хц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   100
         TabIndex        =   59
         Top             =   6000
         Width           =   1600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Рубеж №5 Хц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   100
         TabIndex        =   55
         Top             =   4900
         Width           =   1600
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Рубеж №4 Хц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   100
         TabIndex        =   51
         Top             =   3800
         Width           =   1600
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Рубеж №3 Хц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   47
         Top             =   2700
         Width           =   1600
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Рубеж №2 Хц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   100
         TabIndex        =   43
         Top             =   1600
         Width           =   1600
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2500
         TabIndex        =   40
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Рубеж №1 "
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   38
         Top             =   500
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "X, Y фланги"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7600
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   7815
      Begin VB.TextBox pOtsGr 
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
         Left            =   1800
         TabIndex        =   76
         Text            =   "0"
         Top             =   3500
         Width           =   1000
      End
      Begin VB.CommandButton reshryb 
         BackColor       =   &H00FFC0C0&
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
         Left            =   5500
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6200
         Width           =   1335
      End
      Begin VB.CheckBox p3bat 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   3150
         TabIndex        =   32
         Top             =   6700
         Value           =   1  'Checked
         Width           =   400
      End
      Begin VB.CheckBox p2bat 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   1750
         TabIndex        =   31
         Top             =   6700
         Value           =   1  'Checked
         Width           =   400
      End
      Begin VB.CheckBox p1bat 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   350
         TabIndex        =   30
         Top             =   6700
         Value           =   1  'Checked
         Width           =   400
      End
      Begin VB.OptionButton plev 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   7000
         TabIndex        =   28
         Top             =   3800
         Width           =   400
      End
      Begin VB.OptionButton pzentr 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   7000
         TabIndex        =   27
         Top             =   3100
         Width           =   400
      End
      Begin VB.OptionButton ppr 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   7000
         TabIndex        =   26
         Top             =   2400
         Value           =   -1  'True
         Width           =   400
      End
      Begin VB.TextBox pryb 
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
         Left            =   3800
         TabIndex        =   21
         Text            =   "0"
         Top             =   4600
         Width           =   500
      End
      Begin VB.TextBox pporFr 
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
         Left            =   1600
         TabIndex        =   19
         Text            =   "0"
         Top             =   4600
         Width           =   1000
      End
      Begin VB.TextBox pinter 
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
         Left            =   5600
         TabIndex        =   17
         Text            =   "0"
         Top             =   4600
         Width           =   1000
      End
      Begin VB.TextBox pvFrzo 
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
         Height          =   450
         Left            =   6500
         TabIndex        =   15
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Опредилить фронт ПЗО"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5500
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   500
         Width           =   2000
      End
      Begin VB.TextBox ph 
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
         TabIndex        =   12
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pYp 
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
         Left            =   3500
         TabIndex        =   10
         Text            =   "0"
         Top             =   1700
         Width           =   1500
      End
      Begin VB.TextBox pXp 
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
         Left            =   3500
         TabIndex        =   9
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pYl 
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
         Left            =   600
         TabIndex        =   6
         Text            =   "0"
         Top             =   1700
         Width           =   1500
      End
      Begin VB.TextBox pXl 
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
         Left            =   600
         TabIndex        =   5
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.Label Label38 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Если в составе группы"
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
         TabIndex        =   77
         Top             =   3000
         Width           =   4095
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Отступить от правого края"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   100
         TabIndex        =   75
         Top             =   3500
         Width           =   1500
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 Бат"
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
         Left            =   2900
         TabIndex        =   35
         Top             =   6200
         Width           =   900
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 Бат"
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
         Left            =   1500
         TabIndex        =   34
         Top             =   6200
         Width           =   900
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 Бат"
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
         TabIndex        =   33
         Top             =   6200
         Width           =   900
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Привлечь"
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
         Left            =   1300
         TabIndex        =   29
         Top             =   5500
         Width           =   1600
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Лев край"
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
         Left            =   5500
         TabIndex        =   25
         Top             =   3800
         Width           =   1400
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Центр"
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
         Left            =   5500
         TabIndex        =   24
         Top             =   3100
         Width           =   1000
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Пр край"
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
         Left            =   5500
         TabIndex        =   23
         Top             =   2400
         Width           =   1300
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Отмерять от"
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
         Left            =   5500
         TabIndex        =   22
         Top             =   1700
         Width           =   2000
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "К-во рубежей"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2760
         TabIndex        =   20
         Top             =   4600
         Width           =   1000
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Поражаемый фронт"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   100
         TabIndex        =   18
         Top             =   4600
         Width           =   1500
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Интервал"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4500
         TabIndex        =   16
         Top             =   4600
         Width           =   1000
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фронт"
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
         Left            =   5500
         TabIndex        =   14
         Top             =   1000
         Width           =   1000
      End
      Begin VB.Label Label7 
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
         Left            =   1800
         TabIndex        =   11
         Top             =   2400
         Width           =   400
      End
      Begin VB.Label Label6 
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
         Left            =   3000
         TabIndex        =   8
         Top             =   1700
         Width           =   400
      End
      Begin VB.Label Label5 
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
         Left            =   3000
         TabIndex        =   7
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label Label4 
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
         TabIndex        =   4
         Top             =   1700
         Width           =   400
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Правый край"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   500
         Width           =   2200
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Левый край"
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
         TabIndex        =   1
         Top             =   500
         Width           =   2000
      End
   End
End
Attribute VB_Name = "PZO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X1b As Single, Y1b As Single, X2b As Single, Y2b As Single, X3b As Single, Y3b As Single, Ugolprnalev As Single, Chetych As Single, Shestych As Single, Vtorych As Single, Aatzo As Single, Xp As Single, Yp As Single

Private Sub Command1_Click()
Dim Xl As Single, Yl As Single, Xp As Single, Yp As Single
Xl = pXl: Yl = pYl: Xp = pXp: Yp = pYp
dx = Xl - Xp
dy = Yl - Yp
Pi = 3.14159265358
Dt = Int(Sqr(dx ^ 2 + dy ^ 2) + 0.001)
A1 = Abs(Atn(dy / (dx + 0.001)) / Pi * 30) * 100
If dx > 0 And dy > 0 Then Ygolt = CInt(A1)
If dx < 0 And dy > 0 Then Ygolt = CInt(3000 - A1)
If dx < 0 And dy < 0 Then Ygolt = CInt(3000 + A1)
If dx > 0 And dy < 0 Then Ygolt = CInt(6000 - A1)
Ugolprnalev = Ygolt: Frzo = Dt: pvFrzo = Frzo

End Sub


Private Sub Command2_Click()
PZO.Hide
End Sub

Private Sub navesti1_Click()
If p1bat = 1 Then
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
   xc = X1b: yc = Y1b: hc = ph
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = xc - Xop1
60: dy1 = yc - Yop1
61: dh1 = hc - hop1
   Pi = 3.14159265358
9010: Dt1 = Int(Sqr(dx1 ^ 2 + dy1 ^ 2) + 0.001)
9110: Yr1 = Round(((dh1 + 0.001) / (Dt1 * 0.001 + 0.001)) * 0.95)
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = OZ.pZar1
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar1: vzriv = OZ.pVzr1
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       OZ.poddV0 tz, zar, dv0
              rep1 = OZ.pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
       If rep1 = "Пристрелян" And Dret1 + 2000 < Dt1 Or Dret1 - 2000 > Dt1 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD <= 0 Then
        Dtk = Dt1 - 2000
        Else
        Dtk = Dt1 + 2000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt1 + dXv0c * (v01 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt1 - popvD
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
        Pric1 = Pricisch - Yr1
        Else
        Pric1 = Pricisch + Yr1
       End If
        Yr = Abs(Yr1): Yrr = Yr1: N1 = N: dNtus1 = dNtus
        If snar = "ОФ" And vzriv = "РГМ" Then
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
       daep = kpe * Yr1: preps1 = Round(Pric1 + daep + 0.001)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
If BP.pX1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1: OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1
         OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1: OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
End If
vrv = 0
Else
End If
 ' 2B
 If p2bat = 1 Then
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
         xc = X2b: yc = Y2b: hc = ph
        dx2 = xc - Xop2
104112:  dy2 = yc - Yop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = OZ.pZar2
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar2: vzriv = OZ.pVzr2
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        OZ.poddV0 tz, zar, dv0
               rep2 = OZ.pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
       If rep2 = "Пристрелян" And Dret2 + 2000 < Dt2 Or Dret2 - 2000 > Dt2 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt2 - 1000
        Else
        Dtk = Dt2 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh2 + dXtc * dddt2 + dXv0c * (v02 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt2 - popvD
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
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch - Yr2
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
If BP.pX2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2: OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2
         OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2: OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD)
        OZ.pvDisch2.Text = Int(Disch2): OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
End If
Else
End If
vrv = 0
  '3B
If p3bat = 1 Then
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
          xc = X3b: yc = Y3b: hc = ph
         dx3 = xc - Xop3
1041194:  dy3 = yc - Yop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = OZ.pZar3
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              snar = OZ.pSnar3: vzriv = OZ.pVzr3
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     OZ.poddV0 tz, zar, dv0
            rep3 = OZ.pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
       If rep3 = "Пристрелян" And Dret3 + 2000 < Dt3 Or Dret3 - 2000 > Dt3 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt3 - 1000
        Else
        Dtk = Dt3 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh3 + dXtc * dddt3 + dXv0c * (v03 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt3 - popvD
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
        Pric3 = Pricisch - Yr3
        Else
        Pric3 = Pricisch + Yr3
       End If
        Yr = Abs(Yr3): Yrr = Yr3: N3 = N: dNtus3 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad3 = Ygpad: Ygvozv3 = Ygvozv: Vustra3 = Vustra: ts3 = ts: dXtus3 = dXtus
            Else
            Ygpad3 = Ygpadk: Ygvozv3 = Ygvozvk: Vustra3 = Vustrak: ts3 = tsk: dXtus3 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus3 = 0
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
If BP.pX3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3: OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3
         OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3: OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD)
        OZ.pvDisch3.Text = Int(Disch3): OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
End If
vrv = 0
Else
End If
If Ugolprnalev - 1500 < 0 Then
    Aatzo = Ugolprnalev + 6000 - 1500
    Else
        Aatzo = Ugolprnalev - 1500
End If
If Aatzo + 3000 >= 6000 Then
        prAat = Aatzo + 3000 - 6000
        Else
            prAat = Aatzo + 3000
End If
 If Ygolt1 <= 1500 And prAat >= 4500 Then
    Ygs1 = Abs(Ygolt1 + 6000 - prAat)
    ElseIf Ygolt1 > 4500 And prAat < 1500 Then
        Ygs1 = Abs(Ygolt1 - (prAat + 6000))
        Else
            Ygs1 = Abs(Ygolt1 - prAat)
 End If
 If Ygolt2 <= 1500 And prAat >= 4500 Then
    Ygs2 = Abs(Ygolt2 + 6000 - prAat)
    ElseIf Ygolt2 > 4500 And prAat < 1500 Then
        Ygs2 = Abs(Ygolt2 - (prAat + 6000))
        Else
            Ygs2 = Abs(Ygolt2 - prAat)
 End If
 If Ygolt3 <= 1500 And prAat >= 4500 Then
    Ygs3 = Abs(Ygolt3 + 6000 - prAat)
    ElseIf Ygolt3 > 4500 And prAat < 1500 Then
        Ygs3 = Abs(Ygolt3 - (prAat + 6000))
        Else
            Ygs3 = Abs(Ygolt3 - prAat)
 End If
If p1bat = 1 And p2bat = 1 And p3bat = 1 Then batych = Shestych * 2
If p1bat = 1 And p2bat = 1 And p3bat = 0 Then batych = Chetych * 2
If p1bat = 1 And p2bat = 0 And p3bat = 1 Then batych = Chetych * 2
If p1bat = 0 And p2bat = 1 And p3bat = 1 Then batych = Chetych * 2
If p1bat = 1 And p2bat = 0 And p3bat = 0 Then batych = Vtorych * 2
If p1bat = 0 And p2bat = 1 And p3bat = 0 Then batych = Vtorych * 2
If p1bat = 0 And p2bat = 0 And p3bat = 1 Then batych = Vtorych * 2
pvFr1.Text = Round(batych): pvFr2.Text = Round(batych): pvFr3.Text = Round(batych)
 intve1 = Round(batych / (Dt1 * 0.001 + 0.001) * 0.95)
 intve2 = Round(batych / (Dt2 * 0.001 + 0.001) * 0.95)
 intve3 = Round(batych / (Dt3 * 0.001 + 0.001) * 0.95)
  Sk1 = Round(batych / 4 / (dXtus11 + 0.001)): Sk2 = Round(batych / 4 / (dXtus2 + 0.001)): Sk3 = Round(batych / 4 / (dXtus3 + 0.001))
  If Ygs1 > 750 Then pvtip1.Text = "ФЛАНГ"
       If Ygs1 <= 750 Then pvtip1.Text = "ФРОНТ"
       If Ygs2 > 750 Then pvtip2.Text = "ФЛАНГ"
       If Ygs2 <= 750 Then pvtip2.Text = "ФРОНТ"
       If Ygs3 > 750 Then pvtip3.Text = "ФЛАНГ"
       If Ygs3 <= 750 Then pvtip3.Text = "ФРОНТ"
       If p1bat = 0 Then pvtip1.Text = 0
If p2bat = 0 Then pvtip2.Text = 0
If p3bat = 0 Then pvtip3.Text = 0
       If Ygs1 > 750 Then intve1 = 0
       If Ygs1 <= 750 Then Sk1 = 0
       If Ygs2 > 750 Then intve2 = 0
       If Ygs2 <= 750 Then Sk2 = 0
       If Ygs3 > 750 Then intve3 = 0
       If Ygs3 <= 750 Then Sk3 = 0
 OZ.pvVeer1.Text = intve1: OZ.pvSk1.Text = Sk1
 OZ.pvVeer2.Text = intve2: OZ.pvSk2.Text = Sk2
 OZ.pvVeer3.Text = intve3: OZ.pvSk3.Text = Sk3
OZ.Show
End Sub

Private Sub navesti2_Click()
Dim inetrval As Single
If p1bat = 1 Then
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + X1b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Y1b
hc = ph
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = xc - Xop1
60: dy1 = yc - Yop1
61: dh1 = hc - hop1
   Pi = 3.14159265358
9010: Dt1 = Int(Sqr(dx1 ^ 2 + dy1 ^ 2) + 0.001)
9110: Yr1 = Round(((dh1 + 0.001) / (Dt1 * 0.001 + 0.001)) * 0.95)
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = OZ.pZar1
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar1: vzriv = OZ.pVzr1
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       OZ.poddV0 tz, zar, dv0
              rep1 = OZ.pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
       If rep1 = "Пристрелян" And Dret1 + 2000 < Dt1 Or Dret1 - 2000 > Dt1 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD <= 0 Then
        Dtk = Dt1 - 2000
        Else
        Dtk = Dt1 + 2000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt1 + dXv0c * (v01 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt1 - popvD
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
        Pric1 = Pricisch - Yr1
        Else
        Pric1 = Pricisch + Yr1
       End If
        Yr = Abs(Yr1): Yrr = Yr1: N1 = N: dNtus1 = dNtus
        If snar = "ОФ" And vzriv = "РГМ" Then
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
       daep = kpe * Yr1: preps1 = Round(Pric1 + daep + 0.001)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
If BP.pX1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1: OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1
         OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1: OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
End If
Else
End If
vrv = 0
 ' 2B
 If p2bat = 1 Then
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + X2b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Y2b
hc = ph
        dx2 = xc - Xop2
104112:  dy2 = yc - Yop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = OZ.pZar2
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar2: vzriv = OZ.pVzr2
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        OZ.poddV0 tz, zar, dv0
               rep2 = OZ.pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
       If rep2 = "Пристрелян" And Dret2 + 2000 < Dt2 Or Dret2 - 2000 > Dt2 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt2 - 1000
        Else
        Dtk = Dt2 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh2 + dXtc * dddt2 + dXv0c * (v02 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt2 - popvD
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
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch - Yr2
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
If BP.pX2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2: OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2
         OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2: OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD)
        OZ.pvDisch2.Text = Int(Disch2): OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
End If
Else
End If
vrv = 0
  '3B
  If p3bat = 1 Then
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + X3b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Y3b
hc = ph
         dx3 = xc - Xop3
1041194:  dy3 = yc - Yop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = OZ.pZar3
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              snar = OZ.pSnar3: vzriv = OZ.pVzr3
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     OZ.poddV0 tz, zar, dv0
            rep3 = OZ.pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
       If rep3 = "Пристрелян" And Dret3 + 2000 < Dt3 Or Dret3 - 2000 > Dt3 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt3 - 1000
        Else
        Dtk = Dt3 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh3 + dXtc * dddt3 + dXv0c * (v03 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt3 - popvD
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
        Pric3 = Pricisch - Yr3
        Else
        Pric3 = Pricisch + Yr3
       End If
        Yr = Abs(Yr3): Yrr = Yr3: N3 = N: dNtus3 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad3 = Ygpad: Ygvozv3 = Ygvozv: Vustra3 = Vustra: ts3 = ts: dXtus3 = dXtus
            Else
            Ygpad3 = Ygpadk: Ygvozv3 = Ygvozvk: Vustra3 = Vustrak: ts3 = tsk: dXtus3 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus3 = 0
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
If BP.pX3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3: OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3
         OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3: OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD)
        OZ.pvDisch3.Text = Int(Disch3): OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
End If
vrv = 0
Else
End If
OZ.Show
End Sub

Private Sub navesti3_Click()
Dim inetrval As Single
If p1bat = 1 Then
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 2) + X1b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 2) + Y1b
hc = ph
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = xc - Xop1
60: dy1 = yc - Yop1
61: dh1 = hc - hop1
   Pi = 3.14159265358
9010: Dt1 = Int(Sqr(dx1 ^ 2 + dy1 ^ 2) + 0.001)
9110: Yr1 = Round(((dh1 + 0.001) / (Dt1 * 0.001 + 0.001)) * 0.95)
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = OZ.pZar1
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar1: vzriv = OZ.pVzr1
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       OZ.poddV0 tz, zar, dv0
              rep1 = OZ.pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
       If rep1 = "Пристрелян" And Dret1 + 2000 < Dt1 Or Dret1 - 2000 > Dt1 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD <= 0 Then
        Dtk = Dt1 - 2000
        Else
        Dtk = Dt1 + 2000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt1 + dXv0c * (v01 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt1 - popvD
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
        Pric1 = Pricisch - Yr1
        Else
        Pric1 = Pricisch + Yr1
       End If
        Yr = Abs(Yr1): Yrr = Yr1: N1 = N: dNtus1 = dNtus
        If snar = "ОФ" And vzriv = "РГМ" Then
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
       daep = kpe * Yr1: preps1 = Round(Pric1 + daep + 0.001)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
If BP.pX1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1: OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1
         OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1: OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
End If
Else
End If
vrv = 0
 ' 2B
 If p2bat = 1 Then
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 2) + X2b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 2) + Y2b
hc = ph
        dx2 = xc - Xop2
104112:  dy2 = yc - Yop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = OZ.pZar2
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar2: vzriv = OZ.pVzr2
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        OZ.poddV0 tz, zar, dv0
               rep2 = OZ.pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
       If rep2 = "Пристрелян" And Dret2 + 2000 < Dt2 Or Dret2 - 2000 > Dt2 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt2 - 1000
        Else
        Dtk = Dt2 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh2 + dXtc * dddt2 + dXv0c * (v02 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt2 - popvD
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
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch - Yr2
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
If BP.pX2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2: OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2
         OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2: OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD)
        OZ.pvDisch2.Text = Int(Disch2): OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
End If
Else
End If
vrv = 0
  '3B
  If p3bat = 1 Then
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 2) + X3b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 2) + Y3b
hc = ph
         dx3 = xc - Xop3
1041194:  dy3 = yc - Yop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = OZ.pZar3
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              snar = OZ.pSnar3: vzriv = OZ.pVzr3
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     OZ.poddV0 tz, zar, dv0
            rep3 = OZ.pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
       If rep3 = "Пристрелян" And Dret3 + 2000 < Dt3 Or Dret3 - 2000 > Dt3 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt3 - 1000
        Else
        Dtk = Dt3 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh3 + dXtc * dddt3 + dXv0c * (v03 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt3 - popvD
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
        Pric3 = Pricisch - Yr3
        Else
        Pric3 = Pricisch + Yr3
       End If
        Yr = Abs(Yr3): Yrr = Yr3: N3 = N: dNtus3 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad3 = Ygpad: Ygvozv3 = Ygvozv: Vustra3 = Vustra: ts3 = ts: dXtus3 = dXtus
            Else
            Ygpad3 = Ygpadk: Ygvozv3 = Ygvozvk: Vustra3 = Vustrak: ts3 = tsk: dXtus3 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus3 = 0
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
If BP.pX3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3: OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3
         OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3: OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD)
        OZ.pvDisch3.Text = Int(Disch3): OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
End If
vrv = 0
Else
End If
OZ.Show
End Sub

Private Sub navesti4_Click()
Dim inetrval As Single
If p1bat = 1 Then
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 3) + X1b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 3) + Y1b
hc = ph
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = xc - Xop1
60: dy1 = yc - Yop1
61: dh1 = hc - hop1
   Pi = 3.14159265358
9010: Dt1 = Int(Sqr(dx1 ^ 2 + dy1 ^ 2) + 0.001)
9110: Yr1 = Round(((dh1 + 0.001) / (Dt1 * 0.001 + 0.001)) * 0.95)
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = OZ.pZar1
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar1: vzriv = OZ.pVzr1
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       OZ.poddV0 tz, zar, dv0
              rep1 = OZ.pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
       If rep1 = "Пристрелян" And Dret1 + 2000 < Dt1 Or Dret1 - 2000 > Dt1 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD <= 0 Then
        Dtk = Dt1 - 2000
        Else
        Dtk = Dt1 + 2000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt1 + dXv0c * (v01 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt1 - popvD
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
        Pric1 = Pricisch - Yr1
        Else
        Pric1 = Pricisch + Yr1
       End If
        Yr = Abs(Yr1): Yrr = Yr1: N1 = N: dNtus1 = dNtus
        If snar = "ОФ" And vzriv = "РГМ" Then
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
       daep = kpe * Yr1: preps1 = Round(Pric1 + daep + 0.001)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
If BP.pX1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1: OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1
         OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1: OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
End If
Else
End If
vrv = 0
 ' 2B
 If p2bat = 1 Then
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 3) + X2b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 3) + Y2b
hc = ph
        dx2 = xc - Xop2
104112:  dy2 = yc - Yop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = OZ.pZar2
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar2: vzriv = OZ.pVzr2
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        OZ.poddV0 tz, zar, dv0
               rep2 = OZ.pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
       If rep2 = "Пристрелян" And Dret2 + 2000 < Dt2 Or Dret2 - 2000 > Dt2 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt2 - 1000
        Else
        Dtk = Dt2 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh2 + dXtc * dddt2 + dXv0c * (v02 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt2 - popvD
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
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch - Yr2
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
If BP.pX2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2: OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2
         OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2: OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD)
        OZ.pvDisch2.Text = Int(Disch2): OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
End If
Else
End If
vrv = 0
  '3B
  If p3bat = 1 Then
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 3) + X3b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 3) + Y3b
hc = ph
         dx3 = xc - Xop3
1041194:  dy3 = yc - Yop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = OZ.pZar3
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              snar = OZ.pSnar3: vzriv = OZ.pVzr3
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     OZ.poddV0 tz, zar, dv0
            rep3 = OZ.pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
       If rep3 = "Пристрелян" And Dret3 + 2000 < Dt3 Or Dret3 - 2000 > Dt3 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt3 - 1000
        Else
        Dtk = Dt3 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh3 + dXtc * dddt3 + dXv0c * (v03 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt3 - popvD
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
        Pric3 = Pricisch - Yr3
        Else
        Pric3 = Pricisch + Yr3
       End If
        Yr = Abs(Yr3): Yrr = Yr3: N3 = N: dNtus3 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad3 = Ygpad: Ygvozv3 = Ygvozv: Vustra3 = Vustra: ts3 = ts: dXtus3 = dXtus
            Else
            Ygpad3 = Ygpadk: Ygvozv3 = Ygvozvk: Vustra3 = Vustrak: ts3 = tsk: dXtus3 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus3 = 0
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
If BP.pX3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3: OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3
         OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3: OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD)
        OZ.pvDisch3.Text = Int(Disch3): OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
End If
vrv = 0
Else
End If
OZ.Show
End Sub

Private Sub navesti5_Click()
Dim inetrval As Single
If p1bat = 1 Then
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 4) + X1b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 4) + Y1b
hc = ph
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = xc - Xop1
60: dy1 = yc - Yop1
61: dh1 = hc - hop1
   Pi = 3.14159265358
9010: Dt1 = Int(Sqr(dx1 ^ 2 + dy1 ^ 2) + 0.001)
9110: Yr1 = Round(((dh1 + 0.001) / (Dt1 * 0.001 + 0.001)) * 0.95)
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = OZ.pZar1
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar1: vzriv = OZ.pVzr1
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       OZ.poddV0 tz, zar, dv0
              rep1 = OZ.pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
       If rep1 = "Пристрелян" And Dret1 + 2000 < Dt1 Or Dret1 - 2000 > Dt1 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD <= 0 Then
        Dtk = Dt1 - 2000
        Else
        Dtk = Dt1 + 2000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt1 + dXv0c * (v01 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt1 - popvD
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
        Pric1 = Pricisch - Yr1
        Else
        Pric1 = Pricisch + Yr1
       End If
        Yr = Abs(Yr1): Yrr = Yr1: N1 = N: dNtus1 = dNtus
        If snar = "ОФ" And vzriv = "РГМ" Then
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
       daep = kpe * Yr1: preps1 = Round(Pric1 + daep + 0.001)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
If BP.pX1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1: OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1
         OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1: OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
End If
Else
End If
vrv = 0
 ' 2B
 If p2bat = 1 Then
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 4) + X2b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 4) + Y2b
hc = ph
        dx2 = xc - Xop2
104112:  dy2 = yc - Yop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = OZ.pZar2
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar2: vzriv = OZ.pVzr2
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        OZ.poddV0 tz, zar, dv0
               rep2 = OZ.pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
       If rep2 = "Пристрелян" And Dret2 + 2000 < Dt2 Or Dret2 - 2000 > Dt2 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt2 - 1000
        Else
        Dtk = Dt2 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh2 + dXtc * dddt2 + dXv0c * (v02 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt2 - popvD
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
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch - Yr2
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
If BP.pX2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2: OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2
         OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2: OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD)
        OZ.pvDisch2.Text = Int(Disch2): OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
End If
Else
End If
vrv = 0
  '3B
  If p3bat = 1 Then
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 4) + X3b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 4) + Y3b
hc = ph
         dx3 = xc - Xop3
1041194:  dy3 = yc - Yop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = OZ.pZar3
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              snar = OZ.pSnar3: vzriv = OZ.pVzr3
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     OZ.poddV0 tz, zar, dv0
            rep3 = OZ.pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
       If rep3 = "Пристрелян" And Dret3 + 2000 < Dt3 Or Dret3 - 2000 > Dt3 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt3 - 1000
        Else
        Dtk = Dt3 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh3 + dXtc * dddt3 + dXv0c * (v03 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt3 - popvD
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
        Pric3 = Pricisch - Yr3
        Else
        Pric3 = Pricisch + Yr3
       End If
        Yr = Abs(Yr3): Yrr = Yr3: N3 = N: dNtus3 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad3 = Ygpad: Ygvozv3 = Ygvozv: Vustra3 = Vustra: ts3 = ts: dXtus3 = dXtus
            Else
            Ygpad3 = Ygpadk: Ygvozv3 = Ygvozvk: Vustra3 = Vustrak: ts3 = tsk: dXtus3 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus3 = 0
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
If BP.pX3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3: OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3
         OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3: OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD)
        OZ.pvDisch3.Text = Int(Disch3): OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
End If
vrv = 0
Else
End If
OZ.Show
End Sub

Private Sub navesti6_Click()
Dim inetrval As Single
If p1bat = 1 Then
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 5) + X1b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 5) + Y1b
hc = ph
   Xop1 = BP.pX1: Yop1 = BP.pY1: hop1 = BP.ph1: OH1 = BP.pOH1
   dx1 = xc - Xop1
60: dy1 = yc - Yop1
61: dh1 = hc - hop1
   Pi = 3.14159265358
9010: Dt1 = Int(Sqr(dx1 ^ 2 + dy1 ^ 2) + 0.001)
9110: Yr1 = Round(((dh1 + 0.001) / (Dt1 * 0.001 + 0.001)) * 0.95)
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
       Dt = Dt1: Ygolt = Ygolt1: dh = dh1:   zar = OZ.pZar1
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar1: vzriv = OZ.pVzr1
       dddt1 = dddt: tz = tz1: zc1 = zc
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       OZ.poddV0 tz, zar, dv0
              rep1 = OZ.pRep1: dDov1 = REPER.pvdDov1: Dret1 = REPER.pvDr1: dDr1 = REPER.pvdD1: dN = REPER.pvdN1
       If rep1 = "Пристрелян" And Dret1 + 2000 < Dt1 Or Dret1 - 2000 > Dt1 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD <= 0 Then
        Dtk = Dt1 - 2000
        Else
        Dtk = Dt1 + 2000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh1 + dXtc * dddt1 + dXv0c * (v01 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt1 - popvD
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
        Pric1 = Pricisch - Yr1
        Else
        Pric1 = Pricisch + Yr1
       End If
        Yr = Abs(Yr1): Yrr = Yr1: N1 = N: dNtus1 = dNtus
        If snar = "ОФ" And vzriv = "РГМ" Then
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
       daep = kpe * Yr1: preps1 = Round(Pric1 + daep + 0.001)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
If BP.pX1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1: OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1
         OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1: OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
End If
Else
End If
vrv = 0
 ' 2B
 If p2bat = 1 Then
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 5) + X2b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 5) + Y2b
hc = ph
        dx2 = xc - Xop2
104112:  dy2 = yc - Yop2
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
       Dt = Dt2: Ygolt = Ygolt2: dh = dh2: zar = OZ.pZar2
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       snar = OZ.pSnar2: vzriv = OZ.pVzr2
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
       tz2 = BP.pTz2
        tz = tz2: zc2 = zc
        OZ.poddV0 tz, zar, dv0
               rep2 = OZ.pRep2: dDov2 = REPER.pvdDov2: Dret2 = REPER.pvDr2: dDr2 = REPER.pvdD2: dN = REPER.pvdN2
       If rep2 = "Пристрелян" And Dret2 + 2000 < Dt2 Or Dret2 - 2000 > Dt2 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt2 - 1000
        Else
        Dtk = Dt2 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh2 + dXtc * dddt2 + dXv0c * (v02 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt2 - popvD
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
       Yr = Abs(Yr2): Yrr = Yr2: N2 = N: dNtus2 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad2 = Ygpad: Ygvozv2 = Ygvozv: Vustra2 = Vustra: ts2 = ts: dXtus2 = dXtus
            Else
            Ygpad2 = Ygpadk: Ygvozv2 = Ygvozvk: Vustra2 = Vustrak: ts2 = tsk: dXtus2 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus2 = 0
       If stre = "Мортирная" Then
        Pric2 = Pricisch - Yr2
        Else
        Pric2 = Pricisch + Yr2
       End If
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr2: preps2 = Int(Pric2 + daep)
       End If
       If vzriv = "РГМ" Then dNtus2 = 0
If BP.pX2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2: OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2
         OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2: OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD)
        OZ.pvDisch2.Text = Int(Disch2): OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
End If
Else
End If
vrv = 0
  '3B
  If p3bat = 1 Then
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
Interval = pinter
Pi = 3.14159265358
xc = Cos(Aatzo / 100 * 6 * Pi / 180) * (Interval * 5) + X3b
yc = Sin(Aatzo / 100 * 6 * Pi / 180) * (Interval * 5) + Y3b
hc = ph
         dx3 = xc - Xop3
1041194:  dy3 = yc - Yop3
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
     Dt = Dt3: Ygolt = Ygolt3: dh = dh3:  zar = OZ.pZar3
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
       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              snar = OZ.pSnar3: vzriv = OZ.pVzr3
       If vzriv = "АР-5" Or vzriv = "ДТМ-75" Or vzriv = "В-90" Or vzriv = "Т-90" Then
            tsk = ts: dXtusk = dXtus: Ygvozvk = Ygvozv: Vustrak = Vustra: Ygpadk = Ygpad: Vdk = Vd
            Else
        End If
     tz = BP.pTz3: zc3 = zc
     OZ.poddV0 tz, zar, dv0
            rep3 = OZ.pRep3: dDov3 = REPER.pvdDov3: Dret3 = REPER.pvDr3: dDr3 = REPER.pvdD3: dN = REPER.pvdN3
       If rep3 = "Пристрелян" And Dret3 + 2000 < Dt3 Or Dret3 - 2000 > Dt3 Then soobsch = MsgBox("Дальность переноса выходит за параметры!!!", vbOKOnly, "Предупреждение")
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
        If popvD < 0 Then
        Dtk = Dt3 - 1000
        Else
        Dtk = Dt3 + 1000
        End If
        Dt = Dtk
        If stre = "Мортирная" Then
                OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
                Else
                OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
        End If
        popvdk = dXwc * Wx + dXhc * dhh3 + dXtc * dddt3 + dXv0c * (v03 + dv0)
        popvnapk = dZwc * Wz + zc
       End If
       Dtisch = Dt3 - popvD
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
        Pric3 = Pricisch - Yr3
        Else
        Pric3 = Pricisch + Yr3
       End If
        Yr = Abs(Yr3): Yrr = Yr3: N3 = N: dNtus3 = dNtus
If snar = "ОФ" And vzriv = "РГМ" Then
            Ygpad3 = Ygpad: Ygvozv3 = Ygvozv: Vustra3 = Vustra: ts3 = ts: dXtus3 = dXtus
            Else
            Ygpad3 = Ygpadk: Ygvozv3 = Ygvozvk: Vustra3 = Vustrak: ts3 = tsk: dXtus3 = dXtusk: Vd = Vdk
        End If
       If snar = "ОФ" And vzriv = "АР-5" Then dNtus3 = 0
       If stre = "Мортирная" Then
        OZ.podKPEmort zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 - daep)
       Else
       OZ.podKPE zar, Pricisch, Yrr, kpe
       daep = kpe * Yr3: preps3 = Int(Pric3 + daep)
       End If
       If vzriv = "РГМ" Then dNtus3 = 0
If BP.pX3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3: OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3
         OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3: OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD)
        OZ.pvDisch3.Text = Int(Disch3): OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
End If
vrv = 0
Else
End If
OZ.Show
End Sub

Private Sub pOtsGr_Change()
Dim Xl As Single, Yl As Single, OtsvGr As Single
Xl = pXl: Yl = pYl: Xp = pXp: Yp = pYp
dx = Xl - Xp
dy = Yl - Yp
Pi = 3.14159265358
Dt = Int(Sqr(dx ^ 2 + dy ^ 2) + 0.001)
A1 = Abs(Atn(dy / (dx + 0.001)) / Pi * 30) * 100
If dx > 0 And dy > 0 Then Ygolt = CInt(A1)
If dx < 0 And dy > 0 Then Ygolt = CInt(3000 - A1)
If dx < 0 And dy < 0 Then Ygolt = CInt(3000 + A1)
If dx > 0 And dy < 0 Then Ygolt = CInt(6000 - A1)
Ugolprnalev = Ygolt: Frzo = Dt: pvFrzo = Frzo
OtsvGr = pOtsGr
Xp = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * OtsvGr + Xp
Yp = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * OtsvGr + Yp
End Sub

Private Sub reshryb_Click()
Dim Xl As Single, Yl As Single, Frpor As Single, ryb As Single, Interval As Single, OtsvGr As Single
Interval = pinter
Xl = pXl: Yl = pYl
OtsvGr = pOtsGr
If OtsvGr <> 0 Then
    Xp = Xp: Yp = Yp: Frpor = pporFr
    Else
        Xp = pXp: Yp = pYp: Frpor = pporFr
End If
dx = Xl - Xp
dy = Yl - Yp
Pi = 3.14159265358
Dt = Int(Sqr(dx ^ 2 + dy ^ 2) + 0.001)
A1 = Abs(Atn(dy / (dx + 0.001)) / Pi * 30) * 100
If dx > 0 And dy > 0 Then Ygolt = CInt(A1)
If dx < 0 And dy > 0 Then Ygolt = CInt(3000 - A1)
If dx < 0 And dy < 0 Then Ygolt = CInt(3000 + A1)
If dx > 0 And dy < 0 Then Ygolt = CInt(6000 - A1)
Ugolprnalev = Ygolt: Frzo = Dt
If ppr = True Then
    Ddoze = Frpor / 2
    Xze = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Ddoze + Xp
    Yze = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Ddoze + Yp
    ElseIf pzentr = True Then
        Ddoze = Frzo / 2
        Xze = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Ddoze + Xp
        Yze = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Ddoze + Yp
        ElseIf plev = True Then
            Ddoze = Frpor / 2
            If Ugolprnalev - 3000 < 0 Then
                Ugolprnalev30 = Ugolprnalev + 6000 - 3000
                Else
                    Ugolprnalev30 = Ugolprnalev - 3000
            End If
            Xze = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Ddoze + Xl
            Yze = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Ddoze + Yl
    Else
End If
If Ugolprnalev - 3000 < 0 Then
                Ugolprnalev30 = Ugolprnalev + 6000 - 3000
                Else
                    Ugolprnalev30 = Ugolprnalev - 3000
            End If
pXz1 = Round(Xze): pYz1 = Round(Yze)
ryb = pryb
If Ugolprnalev - 1500 < 0 Then
    Aatzo = Ugolprnalev + 6000 - 1500
    Else
        Aatzo = Ugolprnalev - 1500
End If

If ryb = 2 Then
            Xze2 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze
            Yze2 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze
            pXz2 = Round(Xze2): pYz2 = Round(Yze2)
            ElseIf ryb = 3 Then
            Xze2 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze
            Yze2 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze
            pXz2 = Round(Xze2): pYz2 = Round(Yze2)
            Xze3 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze2
            Yze3 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze2
            pXz3 = Round(Xze3): pYz3 = Round(Yze3)
             ElseIf ryb = 4 Then
            Xze2 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze
            Yze2 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze
            pXz2 = Round(Xze2): pYz2 = Round(Yze2)
            Xze3 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze2
            Yze3 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze2
            pXz3 = Round(Xze3): pYz3 = Round(Yze3)
            Xze4 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze3
            Yze4 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze3
            pXz4 = Round(Xze4): pYz4 = Round(Yze4)
            ElseIf ryb = 5 Then
            Xze2 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze
            Yze2 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze
            pXz2 = Round(Xze2): pYz2 = Round(Yze2)
            Xze3 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze2
            Yze3 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze2
            pXz3 = Round(Xze3): pYz3 = Round(Yze3)
            Xze4 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze3
            Yze4 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze3
            pXz4 = Round(Xze4): pYz4 = Round(Yze4)
            Xze5 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze4
            Yze5 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze4
            pXz5 = Round(Xze5): pYz5 = Round(Yze5)
            ElseIf ryb >= 6 Then
            Xze2 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze
            Yze2 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze
            pXz2 = Round(Xze2): pYz2 = Round(Yze2)
            Xze3 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze2
            Yze3 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze2
            pXz3 = Round(Xze3): pYz3 = Round(Yze3)
            Xze4 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze3
            Yze4 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze3
            pXz4 = Round(Xze4): pYz4 = Round(Yze4)
            Xze5 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze4
            Yze5 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze4
            pXz5 = Round(Xze5): pYz5 = Round(Yze5)
            Xze6 = Cos(Aatzo / 100 * 6 * Pi / 180) * Interval + Xze5
            Yze6 = Sin(Aatzo / 100 * 6 * Pi / 180) * Interval + Yze5
            pXz6 = Round(Xze6): pYz6 = Round(Yze6)
    Else
End If
Chetych = Frpor / 4
Shestych = Frpor / 6
Vtorych = Frpor / 2
If p1bat = 1 And p2bat = 1 And p3bat = 1 Then
    If ppr = True Then
        X1b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Shestych + Xp
        Y1b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Shestych + Yp
        X2b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * (Shestych * 3) + Xp
        Y2b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * (Shestych * 3) + Yp
        X3b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * (Shestych * 5) + Xp
        Y3b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * (Shestych * 5) + Yp
        ElseIf pzentr = True Then
            X2b = Xze: Y2b = Yze
            X1b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 2) + Xze
            Y1b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 2) + Yze
            X3b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * (Shestych * 2) + Xze
            Y3b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * (Shestych * 2) + Yze
            ElseIf plev = True Then
                X3b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Shestych + Xl
                Y3b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Shestych + Yl
                X2b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 3) + Xl
                Y2b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 3) + Yl
                X1b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 5) + Xl
                Y1b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 5) + Yl
        Else
    End If
    ElseIf p1bat = 1 And p2bat = 1 And p3bat = 0 Then
            If ppr = True Then
                X1b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Xp
                Y1b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Yp
                X2b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * (Chetych * 3) + Xp
                Y2b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * (Chetych * 3) + Yp
                ElseIf pzentr = True Then
                        X2b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Xze
                        Y2b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Yze
                        X1b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Xze
                        Y1b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Yze
                    ElseIf plev = True Then
                         X2b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Xl
                        Y2b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Yl
                        X1b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 3) + Xl
                        Y1b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 3) + Yl
                    Else
            End If
        ElseIf p1bat = 1 And p2bat = 0 And p3bat = 1 Then
            If ppr = True Then
                X1b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Xp
                Y1b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Yp
                X3b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * (Chetych * 3) + Xp
                Y3b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * (Chetych * 3) + Yp
                ElseIf pzentr = True Then
                        X3b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Xze
                        Y3b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Yze
                        X1b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Xze
                        Y1b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Yze
                    ElseIf plev = True Then
                         X3b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Xl
                        Y3b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Yl
                        X1b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 3) + Xl
                        Y1b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 3) + Yl
                    Else
            End If
        ElseIf p1bat = 0 And p2bat = 1 And p3bat = 1 Then
            If ppr = True Then
                X2b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Xp
                Y2b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Yp
                X3b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * (Chetych * 3) + Xp
                Y3b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * (Chetych * 3) + Yp
                ElseIf pzentr = True Then
                        X3b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Xze
                        Y3b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Chetych + Yze
                        X2b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Xze
                        Y2b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Yze
                    ElseIf plev = True Then
                         X3b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Xl
                        Y3b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Chetych + Yl
                        X2b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 3) + Xl
                        Y2b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * (Shestych * 3) + Yl
                    Else
            End If
        ElseIf p1bat = 1 And p2bat = 0 And p3bat = 0 Then
               If ppr = True Then
                X1b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Vtorych + Xp
                Y1b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Vtorych + Yp
                ElseIf pzentr = True Then
                        X1b = Xze
                        Y1b = Yze
                    ElseIf plev = True Then
                         X1b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Vtorych + Xl
                        Y1b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Vtorych + Yl
                    Else
            End If
        ElseIf p1bat = 0 And p2bat = 1 And p3bat = 0 Then
                If ppr = True Then
                X2b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Vtorych + Xp
                Y2b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Vtorych + Yp
                ElseIf pzentr = True Then
                        X2b = Xze
                        Y2b = Yze
                    ElseIf plev = True Then
                         X2b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Vtorych + Xl
                        Y2b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Vtorych + Yl
                    Else
            End If
        ElseIf p1bat = 0 And p2bat = 0 And p3bat = 1 Then
                 If ppr = True Then
                X3b = Cos(Ugolprnalev / 100 * 6 * Pi / 180) * Vtorych + Xp
                Y3b = Sin(Ugolprnalev / 100 * 6 * Pi / 180) * Vtorych + Yp
                ElseIf pzentr = True Then
                        X3b = Xze
                        Y3b = Yze
                    ElseIf plev = True Then
                         X3b = Cos(Ugolprnalev30 / 100 * 6 * Pi / 180) * Vtorych + Xl
                        Y3b = Sin(Ugolprnalev30 / 100 * 6 * Pi / 180) * Vtorych + Yl
                    Else
            End If
    Else
End If

End Sub
