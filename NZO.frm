VERSION 5.00
Begin VB.Form NZO 
   Caption         =   "НЗО"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   18120
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3 Батарея участок"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4000
      Left            =   8500
      TabIndex        =   63
      Top             =   4500
      Width           =   4000
      Begin VB.TextBox pvhc3 
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
         Left            =   1100
         TabIndex        =   73
         Text            =   "0"
         Top             =   3300
         Width           =   1000
      End
      Begin VB.TextBox pvYce3 
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
         Left            =   1100
         TabIndex        =   72
         Text            =   "0"
         Top             =   2600
         Width           =   1500
      End
      Begin VB.TextBox pvXce3 
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
         Left            =   1100
         TabIndex        =   71
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.TextBox pv3BFr 
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
         Left            =   1100
         TabIndex        =   70
         Text            =   "0"
         Top             =   1100
         Width           =   1000
      End
      Begin VB.TextBox pvZO3 
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
         Left            =   1100
         TabIndex        =   69
         Text            =   "0"
         Top             =   400
         Width           =   2300
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   68
         Top             =   3300
         Width           =   500
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   67
         Top             =   2600
         Width           =   500
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   66
         Top             =   1900
         Width           =   500
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фронт"
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
         Left            =   100
         TabIndex        =   65
         Top             =   1100
         Width           =   900
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "НЗО"
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
         Left            =   100
         TabIndex        =   64
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2 Батарея участок"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4000
      Left            =   4300
      TabIndex        =   52
      Top             =   4500
      Width           =   4000
      Begin VB.TextBox pvhc2 
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
         Left            =   1100
         TabIndex        =   62
         Text            =   "0"
         Top             =   3300
         Width           =   1000
      End
      Begin VB.TextBox pvYce2 
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
         Left            =   1100
         TabIndex        =   61
         Text            =   "0"
         Top             =   2600
         Width           =   1500
      End
      Begin VB.TextBox pvXce2 
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
         Left            =   1100
         TabIndex        =   60
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.TextBox pv2BFr 
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
         Left            =   1100
         TabIndex        =   59
         Text            =   "0"
         Top             =   1100
         Width           =   1000
      End
      Begin VB.TextBox pvZO2 
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
         Left            =   1100
         TabIndex        =   58
         Text            =   "0"
         Top             =   400
         Width           =   2300
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   57
         Top             =   3300
         Width           =   500
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   56
         Top             =   2600
         Width           =   500
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   55
         Top             =   1900
         Width           =   500
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фронт"
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
         Left            =   100
         TabIndex        =   54
         Top             =   1100
         Width           =   900
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "НЗО"
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
         Left            =   100
         TabIndex        =   53
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1 Батарея участок"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4000
      Left            =   100
      TabIndex        =   41
      Top             =   4500
      Width           =   4000
      Begin VB.TextBox pvhc1 
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
         Left            =   1100
         TabIndex        =   51
         Text            =   "0"
         Top             =   3300
         Width           =   1000
      End
      Begin VB.TextBox pvYce1 
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
         Left            =   1100
         TabIndex        =   50
         Text            =   "0"
         Top             =   2600
         Width           =   1500
      End
      Begin VB.TextBox pvXce1 
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
         Left            =   1100
         TabIndex        =   49
         Text            =   "0"
         Top             =   1900
         Width           =   1500
      End
      Begin VB.TextBox pv1BFr 
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
         Left            =   1100
         TabIndex        =   48
         Text            =   "0"
         Top             =   1100
         Width           =   1000
      End
      Begin VB.TextBox pvZO1 
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
         Left            =   1100
         TabIndex        =   47
         Text            =   "0"
         Top             =   400
         Width           =   2300
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   46
         Top             =   3300
         Width           =   500
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   45
         Top             =   2600
         Width           =   500
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   44
         Top             =   1900
         Width           =   500
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фронт"
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
         Left            =   100
         TabIndex        =   43
         Top             =   1100
         Width           =   900
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "НЗО"
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
         Left            =   100
         TabIndex        =   42
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Привлечь"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4300
      Left            =   14800
      TabIndex        =   30
      Top             =   100
      Width           =   5400
      Begin VB.TextBox pvFrontReal 
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
         Left            =   2200
         TabIndex        =   40
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pFrontMax 
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
         Left            =   2200
         TabIndex        =   38
         Text            =   "0"
         Top             =   1900
         Width           =   1000
      End
      Begin VB.CheckBox p3Bat 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check3"
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
         Left            =   4100
         TabIndex        =   33
         Top             =   1100
         Width           =   300
      End
      Begin VB.CheckBox p2Bat 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check2"
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
         Left            =   2300
         TabIndex        =   32
         Top             =   1100
         Width           =   300
      End
      Begin VB.CheckBox p1Bat 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
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
         Left            =   500
         TabIndex        =   31
         Top             =   1100
         Width           =   300
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Реальные размеры НЗО"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   200
         TabIndex        =   39
         Top             =   2700
         Width           =   1695
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Максимальные размеры НЗО"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   200
         TabIndex        =   37
         Top             =   1800
         Width           =   2000
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  3 Бат"
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
         Left            =   3800
         TabIndex        =   36
         Top             =   400
         Width           =   1000
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  2 Бат"
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
         Left            =   2000
         TabIndex        =   35
         Top             =   400
         Width           =   1000
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  1 Бат"
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
         Left            =   200
         TabIndex        =   34
         Top             =   400
         Width           =   1000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Фланги А, Д"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4300
      Left            =   7450
      TabIndex        =   14
      Top             =   100
      Width           =   7200
      Begin VB.TextBox pMcp 
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
         Left            =   3600
         TabIndex        =   76
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.CommandButton nzoAD 
         BackColor       =   &H00FFC0C0&
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
         Height          =   1100
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1800
         Width           =   1200
      End
      Begin VB.TextBox pDp 
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
         Left            =   3600
         TabIndex        =   28
         Text            =   "0"
         Top             =   2600
         Width           =   1500
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
         Height          =   450
         Left            =   3600
         TabIndex        =   27
         Text            =   "0"
         Top             =   1800
         Width           =   1500
      End
      Begin VB.TextBox pMcl 
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
         Left            =   700
         TabIndex        =   26
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.TextBox pDl 
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
         Left            =   700
         TabIndex        =   25
         Text            =   "0"
         Top             =   2600
         Width           =   1500
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
         Height          =   450
         Left            =   700
         TabIndex        =   24
         Text            =   "0"
         Top             =   1800
         Width           =   1500
      End
      Begin VB.ComboBox pNKP 
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
         Left            =   2300
         TabIndex        =   16
         Text            =   "1"
         Top             =   500
         Width           =   1000
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мс="
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
         Left            =   3000
         TabIndex        =   75
         Top             =   3400
         Width           =   500
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   3000
         TabIndex        =   23
         Top             =   2600
         Width           =   500
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   22
         Top             =   1800
         Width           =   500
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мс="
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
         Left            =   100
         TabIndex        =   21
         Top             =   3400
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   20
         Top             =   2600
         Width           =   500
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   3000
         TabIndex        =   19
         Top             =   1800
         Width           =   500
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "     Правая"
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
         Left            =   3000
         TabIndex        =   18
         Top             =   1200
         Width           =   2100
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "          Левая"
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
         Left            =   100
         TabIndex        =   17
         Top             =   1200
         Width           =   2100
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ КП="
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
         Left            =   1320
         TabIndex        =   15
         Top             =   500
         Width           =   1000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Фланги Х, У"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4300
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   7200
      Begin VB.ComboBox pvplNZO 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4500
         TabIndex        =   78
         Text            =   "0"
         Top             =   3200
         Width           =   2200
      End
      Begin VB.CommandButton nzoXY 
         BackColor       =   &H00FFC0C0&
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
         Height          =   1100
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1000
         Width           =   1200
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
         Height          =   450
         Left            =   3500
         TabIndex        =   12
         Text            =   "0"
         Top             =   1800
         Width           =   1500
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
         Height          =   450
         Left            =   3500
         TabIndex        =   11
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox phc 
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
         Left            =   2400
         TabIndex        =   8
         Text            =   "0"
         Top             =   2600
         Width           =   1000
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
         Height          =   450
         Left            =   600
         TabIndex        =   7
         Text            =   "0"
         Top             =   1800
         Width           =   1500
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
         Height          =   450
         Left            =   600
         TabIndex        =   6
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Плановый НЗО"
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
         Left            =   4500
         TabIndex        =   77
         Top             =   2600
         Width           =   2200
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   3000
         TabIndex        =   10
         Top             =   1800
         Width           =   500
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   3000
         TabIndex        =   9
         Top             =   1000
         Width           =   500
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   2000
         TabIndex        =   5
         Top             =   2600
         Width           =   500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   4
         Top             =   1800
         Width           =   500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
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
         Height          =   400
         Left            =   100
         TabIndex        =   3
         Top             =   1000
         Width           =   500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "         Правая"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   500
         Width           =   2000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "         Левая"
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
         Left            =   100
         TabIndex        =   1
         Top             =   500
         Width           =   2000
      End
   End
End
Attribute VB_Name = "NZO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NZO.Hide
End Sub
Private Sub nzoAD_Click()
Dim Front As Single, Xpr As Single, Ypr As Single, Xlev As Single, Ylev As Single, hc As Single, ktostr As Single
Dim o1Bat As String, o2Bat As String, o3Bat As String
  Dim dDov1 As Single, Dret1 As Single, dDr1 As Single, dN As Single, zo11 As Single, nkp As Single
 Dim rep1 As String, rep2 As String, rep3 As String
 Dim Apr As Single, Dprav As Single, Alev As Single, Dlev As Single, Mcpr As Single, Mclev As Single
 Dim Xkp As Single, Ykp As Single, hkp As Single
230 vvz = 0: hc = 0
Front = pFrontMax
o1Bat = p1bat: o2Bat = p2bat: o3Bat = p3bat
If o1Bat = True And o2Bat = True And o3Bat = True Then
    ktostr = 1
    ElseIf o1Bat = True And o2Bat = True And o3Bat = False Then
        ktostr = 2
        ElseIf o1Bat = True And o2Bat = False And o3Bat = True Then
            ktostr = 3
            ElseIf o1Bat = False And o2Bat = True And o3Bat = True Then
                ktostr = 4
                ElseIf o1Bat = True And o2Bat = False And o3Bat = False Then
                    ktostr = 5
                    ElseIf o1Bat = False And o2Bat = True And o3Bat = False Then
                         ktostr = 6
                         ElseIf o1Bat = False And o2Bat = False And o3Bat = True Then
                            ktostr = 7
                            Else
                                    ktostr = 1
End If
nkp = pNKp: Apr = pAp: Dprav = pDp: Alev = pAl: Dlev = pDl: Mcpr = pMcp: Mclev = pMcl
   If nkp = 1 Then Xkp = BP.pXkp1: Ykp = BP.pYkp1: hkp = BP.phkp1
   If nkp = 2 Then Xkp = BP.pXkp2: Ykp = BP.pYkp2: hkp = BP.phkp2
   If nkp = 3 Then Xkp = BP.pXkp3: Ykp = BP.pYkp3: hkp = BP.phkp3
   If nkp = 4 Then Xkp = BP.pXkp4: Ykp = BP.pYkp4: hkp = BP.phkp4
   If nkp = 5 Then Xkp = BP.pXkp5: Ykp = BP.pYkp5: hkp = BP.phkp5
      Xpr = Cos(Apr / 100 * 6 * 3.141592 / 180) * Dprav + Xkp
      Ypr = Sin(Apr / 100 * 6 * 3.141592 / 180) * Dprav + Ykp
      hpr = Mcpr * (Dprav * 0.001) * 1.05 + hkp
      Xlev = Cos(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Xkp
      Ylev = Sin(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Ykp
      hlev = Mclev * (Dlev * 0.001) * 1.05 + hkp
       hc = (hpr - hlev) / 2 + hlev
'''Flangi
240 dxzo = Xlev - Xpr: dyzo = Ylev - Ypr
    Frzo = Sqr(dxzo ^ 2 + dyzo ^ 2)
    Azo = Abs(Atn(dyzo / (dxzo + 0.001)) / 3.141592 * 30) * 100
  If dxzo > 0 And dyzo > 0 Then Ygolzo = Int(Azo)
  If dxzo < 0 And dyzo > 0 Then Ygolzo = Int(3000 - Azo)
  If dxzo < 0 And dyzo < 0 Then Ygolzo = Int(3000 + Azo)
  If dxzo > 0 And dyzo < 0 Then Ygolzo = Int(6000 - Azo)
 
  If ktostr = 1 And Frzo < Front Then dzo = Frzo / 6
  If ktostr = 1 And Frzo > Front Then dzo = Front / 6
  If ktostr = 2 Or ktostr = 3 Or ktostr = 4 And Frzo < Front Then dzo = Frzo / 4
  If ktostr = 2 Or ktostr = 3 Or ktostr = 4 And Frzo > Front Then dzo = Front / 4
  If ktostr = 5 Or ktostr = 6 Or ktostr = 7 And Frzo < Front Then dzo = Frzo / 2
  If ktostr = 5 Or ktostr = 6 Or ktostr = 7 And Frzo > Front Then dzo = Front / 2
240101
'1B
   Xzzo1 = Cos(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo + Xpr
   Yzzo1 = Sin(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo + Ypr
'2B
   Xzzo2 = Cos(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo * 3 + Xpr
   Yzzo2 = Sin(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo * 3 + Ypr
'3B
   Xzzo3 = Cos(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo * 5 + Xpr
   Yzzo3 = Sin(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo * 5 + Ypr
   
 If ktostr = 3 Then Xzzo3 = Xzzo2: Yzzo3 = Yzzo2: Xzzo2 = 0: Yzzo2 = 0
 If ktostr = 4 Then Xzzo3 = Xzzo2: Yzzo3 = Yzzo2: Xzzo2 = Xzzo1: Yzzo2 = Yzzo1: Xzzo1 = 0: Yzzo1 = 0
 If ktostr = 6 Then Xzzo2 = Xzzo1: Yzzo2 = Yzzo1: Xzzo1 = 0: Yzzo1 = 0: Xzzo3 = 0: Yzzo3 = 0
 If ktostr = 5 Then Xzzo2 = 0: Yzzo2 = 0: Xzzo3 = 0: Yzzo3 = 0
 If ktostr = 7 Then Xzzo3 = Xzzo1: Yzzo3 = Yzzo1: Xzzo1 = 0: Yzzo1 = 0: Xzzo2 = 0: Yzzo2 = 0
24011  Xc1 = Xzzo1: Yc1 = Yzzo1: Xc2 = Xzzo2: Yc2 = Yzzo2: Xc3 = Xzzo3: Yc3 = Yzzo3
If o1Bat = True Then
    pvXce1.Text = Round(Xc1): pvYce1.Text = Round(Yc1): pvhc1.Text = Round(hc)
    Else
        pvXce1.Text = 0: pvYce1.Text = 0: pvhc1.Text = 0
End If
If o2Bat = True Then
    pvXce2.Text = Round(Xc2): pvYce2.Text = Round(Yc2): pvhc2.Text = Round(hc)
    Else
        pvXce2.Text = 0: pvYce2.Text = 0: pvhc2.Text = 0
End If
If o3Bat = True Then
    pvXce3.Text = Round(Xc3): pvYce3.Text = Round(Yc3): pvhc3.Text = Round(hc)
    Else
        pvXce3.Text = 0: pvYce3.Text = 0: pvhc3.Text = 0
End If
 If Frzo > Front Then
 Front = Front
 ElseIf Frzo < Front And Frzo > 0 Then
 Front = Frzo
 Else
 Front = Front
 End If
 If ktostr = 1 Then batych = Front / 3
 If ktostr > 1 And ktostr < 5 Then batych = Front / 2
 If ktostr > 4 Then batych = Front
 If ktostr = 1 Or ktostr = 2 Or ktostr = 3 Or ktostr = 5 Then batych1 = batych
 If ktostr = 1 Or ktostr = 2 Or ktostr = 4 Or ktostr = 6 Then batych2 = batych
 If ktostr = 1 Or ktostr = 3 Or ktostr = 4 Or ktostr = 7 Then batych3 = batych
 
 pvFrontReal.Text = Round(Frzo)
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
zo11 = 1
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
   If zo11 = 1 Then
   xc = Xc1: yc = Yc1: hc = hc
   Else
   xc = pXc: yc = pYc: hc = phc
   End If
   Xc1 = xc: Yc1 = yc: hc1 = hc
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

snar = OZ.pSnar1: vzriv = OZ.pVzr1

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
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
       daep = kpe * Yr1: preps1 = Round(Pric1 + daep + 0.001)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
        Xc1 = xc: Yc1 = yc: hc1 = hc
If BP.pX1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1: OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1
         OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1: OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
        OZ.pvSnar1.Text = 0: OZ.pvvzr1.Text = 0: OZ.pvZar1.Text = 0: OZ.pvPric1.Text = 0: OZ.pvN1.Text = 0: OZ.pvDov1.Text = 0
         OZ.pvdXtus1.Text = 0: OZ.pvdNtus1.Text = 0: OZ.pvPolet1.Text = 0: OZ.pvVustra1.Text = 0
        OZ.pvVd1.Text = 0: OZ.pvDt1.Text = 0: OZ.pvYgt1.Text = 0: OZ.pvDovt1.Text = 0: OZ.pvYr1.Text = 0: OZ.pvOH1.Text = 0: OZ.pvdD1.Text = 0
        OZ.pvDisch1.Text = 0: OZ.pvdDov1.Text = 0
End If
vrv = 0
 ' 2B
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
        If zo11 = 1 Then
         xc = Xc2: yc = Yc2: hc = hc
         Else
         xc = xc: yc = yc: hc = hc
         End If
         Xc2 = xc: Yc2 = yc: hc2 = hc
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
     
snar = OZ.pSnar2: vzriv = OZ.pVzr2

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
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
If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "РГМ" Then
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
      Xc2 = xc: Yc2 = yc: hc2 = hc
If BP.pX2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2: OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2
         OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2: OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD)
        OZ.pvDisch2.Text = Int(Disch2): OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
              OZ.pvSnar2.Text = 0: OZ.pvvzr2.Text = 0: OZ.pvZar2.Text = 0: OZ.pvPric2.Text = 0: OZ.pvN2.Text = 0: OZ.pvDov2.Text = 0
         OZ.pvdXtus2.Text = 0: OZ.pvdNtus2.Text = 0: OZ.pvPolet2.Text = 0: OZ.pvVustra2.Text = 0
        OZ.pvVd2.Text = 0: OZ.pvDt2.Text = 0: OZ.pvYgt2.Text = 0: OZ.pvDovt2.Text = 0: OZ.pvYr2.Text = 0: OZ.pvOH2.Text = 0: OZ.pvdD2.Text = 0
        OZ.pvDisch2.Text = 0: OZ.pvdDov2.Text = 0
End If
vrv = 0
  '3B
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
        If zo11 = 1 Then
          xc = Xc3: yc = Yc3: hc = hc
          Else
          xc = xc: yc = yc: hc = hc
          End If
          Xc3 = xc: Yc3 = yc: hc3 = hc
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
       
snar = OZ.pSnar3: vzriv = OZ.pVzr3

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              
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
If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "РГМ" Then
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
       Xc3 = xc: Yc3 = yc: hc3 = hc
If BP.pX3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3: OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3
         OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3: OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD)
        OZ.pvDisch3.Text = Int(Disch3): OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
                     OZ.pvSnar3.Text = 0: OZ.pvvzr3.Text = 0: OZ.pvZar3.Text = 0: OZ.pvPric3.Text = 0: OZ.pvN3.Text = 0: OZ.pvDov3.Text = 0
         OZ.pvdXtus3.Text = 0: OZ.pvdNtus3.Text = 0: OZ.pvPolet3.Text = 0: OZ.pvVustra3.Text = 0
        OZ.pvVd3.Text = 0: OZ.pvDt3.Text = 0: OZ.pvYgt3.Text = 0: OZ.pvDovt3.Text = 0: OZ.pvYr3.Text = 0: OZ.pvOH3.Text = 0: OZ.pvdD3.Text = 0
        OZ.pvDisch3.Text = 0: OZ.pvdDov3.Text = 0
End If
vrv = 0
If Ygolzo - 1500 <= 0 Then
        Aatzo = Ygolzo + 6000 - 1500
        Else
            Aatzo = Ygolzo - 1500
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
pv1BFr.Text = Round(batych1): pv2BFr.Text = Round(batych2): pv3BFr.Text = Round(batych3)
 intve1 = Round(batych1 / (Dt1 * 0.001 + 0.001) * 0.95)
 intve2 = Round(batych2 / (Dt2 * 0.001 + 0.001) * 0.95)
 intve3 = Round(batych3 / (Dt3 * 0.001 + 0.001) * 0.95)
  Sk1 = Round(batych1 / 4 / (dXtus11 + 0.001)): Sk2 = Round(batych2 / 4 / (dXtus2 + 0.001)): Sk3 = Round(batych3 / 4 / (dXtus3 + 0.001))
  If Ygs1 > 750 Then pvZO1.Text = "ФЛАНГОВЫЙ"
       If Ygs1 <= 750 Then pvZO1.Text = "ФРОНТАЛЬНЫЙ"
       If Ygs2 > 750 Then pvZO2.Text = "ФЛАНГОВЫЙ"
       If Ygs2 <= 750 Then pvZO2.Text = "ФРОНТАЛЬНЫЙ"
       If Ygs3 > 750 Then pvZO3.Text = "ФЛАНГОВЫЙ"
       If Ygs3 <= 750 Then pvZO3.Text = "ФРОНТАЛЬНЫЙ"
If p1bat = False Then pvZO1.Text = 0
If p2bat = False Then pvZO2.Text = 0
If p3bat = False Then pvZO3.Text = 0
       If Ygs1 > 750 Then intve1 = 0
       If Ygs1 <= 750 Then Sk1 = 0
       If Ygs2 > 750 Then intve2 = 0
       If Ygs2 <= 750 Then Sk2 = 0
       If Ygs3 > 750 Then intve3 = 0
       If Ygs3 <= 750 Then Sk3 = 0
 OZ.pvVeer1.Text = intve1: OZ.pvSk1.Text = Sk1
 OZ.pvVeer2.Text = intve2: OZ.pvSk2.Text = Sk2
 OZ.pvVeer3.Text = intve3: OZ.pvSk3.Text = Sk3

End Sub

Private Sub nzoXY_Click()
Dim Front As Single, Xpr As Single, Ypr As Single, Xlev As Single, Ylev As Single, hc As Single, ktostr As Single
Dim o1Bat As String, o2Bat As String, o3Bat As String
  Dim dDov1 As Single, Dret1 As Single, dDr1 As Single, dN As Single, zo11 As Single
 Dim rep1 As String, rep2 As String, rep3 As String
230 vvz = 0: hc = 0
Front = pFrontMax
o1Bat = p1bat: o2Bat = p2bat: o3Bat = p3bat
If o1Bat = True And o2Bat = True And o3Bat = True Then
    ktostr = 1
    ElseIf o1Bat = True And o2Bat = True And o3Bat = False Then
        ktostr = 2
        ElseIf o1Bat = True And o2Bat = False And o3Bat = True Then
            ktostr = 3
            ElseIf o1Bat = False And o2Bat = True And o3Bat = True Then
                ktostr = 4
                ElseIf o1Bat = True And o2Bat = False And o3Bat = False Then
                    ktostr = 5
                    ElseIf o1Bat = False And o2Bat = True And o3Bat = False Then
                         ktostr = 6
                         ElseIf o1Bat = False And o2Bat = False And o3Bat = True Then
                            ktostr = 7
                            Else
                                    ktostr = 1
End If
Xpr = pXp: Ypr = pYp: Xlev = pXl: Ylev = pYl: hc = phc
'''Flangi
240 dxzo = Xlev - Xpr: dyzo = Ylev - Ypr
    Frzo = Sqr(dxzo ^ 2 + dyzo ^ 2)
    Azo = Abs(Atn(dyzo / (dxzo + 0.001)) / 3.141592 * 30) * 100
  If dxzo > 0 And dyzo > 0 Then Ygolzo = Int(Azo)
  If dxzo < 0 And dyzo > 0 Then Ygolzo = Int(3000 - Azo)
  If dxzo < 0 And dyzo < 0 Then Ygolzo = Int(3000 + Azo)
  If dxzo > 0 And dyzo < 0 Then Ygolzo = Int(6000 - Azo)
 
  If ktostr = 1 And Frzo < Front Then dzo = Frzo / 6
  If ktostr = 1 And Frzo > Front Then dzo = Front / 6
  If ktostr = 2 Or ktostr = 3 Or ktostr = 4 And Frzo < Front Then dzo = Frzo / 4
  If ktostr = 2 Or ktostr = 3 Or ktostr = 4 And Frzo > Front Then dzo = Front / 4
  If ktostr = 5 Or ktostr = 6 Or ktostr = 7 And Frzo < Front Then dzo = Frzo / 2
  If ktostr = 5 Or ktostr = 6 Or ktostr = 7 And Frzo > Front Then dzo = Front / 2
240101
'1B
   Xzzo1 = Cos(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo + Xpr
   Yzzo1 = Sin(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo + Ypr
'2B
   Xzzo2 = Cos(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo * 3 + Xpr
   Yzzo2 = Sin(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo * 3 + Ypr
'3B
   Xzzo3 = Cos(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo * 5 + Xpr
   Yzzo3 = Sin(Ygolzo / 100 * 6 * 3.141592 / 180) * dzo * 5 + Ypr
   
 If ktostr = 3 Then Xzzo3 = Xzzo2: Yzzo3 = Yzzo2: Xzzo2 = 0: Yzzo2 = 0
 If ktostr = 4 Then Xzzo3 = Xzzo2: Yzzo3 = Yzzo2: Xzzo2 = Xzzo1: Yzzo2 = Yzzo1: Xzzo1 = 0: Yzzo1 = 0
 If ktostr = 6 Then Xzzo2 = Xzzo1: Yzzo2 = Yzzo1: Xzzo1 = 0: Yzzo1 = 0: Xzzo3 = 0: Yzzo3 = 0
 If ktostr = 5 Then Xzzo2 = 0: Yzzo2 = 0: Xzzo3 = 0: Yzzo3 = 0
 If ktostr = 7 Then Xzzo3 = Xzzo1: Yzzo3 = Yzzo1: Xzzo1 = 0: Yzzo1 = 0: Xzzo2 = 0: Yzzo2 = 0
24011  Xc1 = Xzzo1: Yc1 = Yzzo1: Xc2 = Xzzo2: Yc2 = Yzzo2: Xc3 = Xzzo3: Yc3 = Yzzo3
If o1Bat = True Then
    pvXce1.Text = Round(Xc1): pvYce1.Text = Round(Yc1): pvhc1.Text = phc
    Else
        pvXce1.Text = 0: pvYce1.Text = 0: pvhc1.Text = 0
End If
If o2Bat = True Then
    pvXce2.Text = Round(Xc2): pvYce2.Text = Round(Yc2): pvhc2.Text = phc
    Else
        pvXce2.Text = 0: pvYce2.Text = 0: pvhc2.Text = 0
End If
If o3Bat = True Then
    pvXce3.Text = Round(Xc3): pvYce3.Text = Round(Yc3): pvhc3.Text = phc
    Else
        pvXce3.Text = 0: pvYce3.Text = 0: pvhc3.Text = 0
End If
 If Frzo > Front Then
 Front = Front
 ElseIf Frzo < Front And Frzo > 0 Then
 Front = Frzo
 Else
 Front = Front
 End If
 If ktostr = 1 Then batych = Front / 3
 If ktostr > 1 And ktostr < 5 Then batych = Front / 2
 If ktostr > 4 Then batych = Front
 If ktostr = 1 Or ktostr = 2 Or ktostr = 3 Or ktostr = 5 Then batych1 = batych
 If ktostr = 1 Or ktostr = 2 Or ktostr = 4 Or ktostr = 6 Then batych2 = batych
 If ktostr = 1 Or ktostr = 3 Or ktostr = 4 Or ktostr = 7 Then batych3 = batych
 
 pvFrontReal.Text = Round(Frzo)
''''''''''''''''''''''''''''''''''OGNEVUE podprogr'''''''''''''''''''''
      '1B
zo11 = 1
ras = 0: h = BP.ph: hop1 = BP.ph1: tz1 = BP.pTz1: hmet = BP.phmet: stre = OZ.pStre1
If h = 0 Then h = 750
215: dhh1 = (h - 750) + ((hmet - hop1) / 10)
   If zo11 = 1 Then
   xc = Xc1: yc = Yc1: hc = hc
   Else
   xc = pXc: yc = pYc: hc = phc
   End If
   Xc1 = xc: Yc1 = yc: hc1 = hc
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

snar = OZ.pSnar1: vzriv = OZ.pVzr1

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
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
       daep = kpe * Yr1: preps1 = Round(Pric1 + daep + 0.001)
       End If
       If vzriv = "РГМ" Then dNtus1 = 0
        Xc1 = xc: Yc1 = yc: hc1 = hc
If BP.pX1 <> 0 Then
        OZ.pvSnar1.Text = snar: OZ.pvvzr1.Text = vzriv: OZ.pvZar1.Text = zar: OZ.pvPric1.Text = preps1: OZ.pvN1.Text = CInt(N1): OZ.pvDov1.Text = dovisch1
         OZ.pvdXtus1.Text = dXtus11: OZ.pvdNtus1.Text = dNtus1: OZ.pvPolet1.Text = ts1: OZ.pvVustra1.Text = Vustra1
        OZ.pvVd1.Text = Vd: OZ.pvDt1.Text = Dt1: OZ.pvYgt1.Text = Ygolt1: OZ.pvDovt1.Text = Dovort1: OZ.pvYr1.Text = Yr1: OZ.pvOH1.Text = OH1: OZ.pvdD1.Text = CInt(popvD)
        OZ.pvDisch1.Text = Int(Disch1): OZ.pvdDov1.Text = CInt(popvnap1)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "1 Батарея")
Else
        OZ.pvSnar1.Text = 0: OZ.pvvzr1.Text = 0: OZ.pvZar1.Text = 0: OZ.pvPric1.Text = 0: OZ.pvN1.Text = 0: OZ.pvDov1.Text = 0
         OZ.pvdXtus1.Text = 0: OZ.pvdNtus1.Text = 0: OZ.pvPolet1.Text = 0: OZ.pvVustra1.Text = 0
        OZ.pvVd1.Text = 0: OZ.pvDt1.Text = 0: OZ.pvYgt1.Text = 0: OZ.pvDovt1.Text = 0: OZ.pvYr1.Text = 0: OZ.pvOH1.Text = 0: OZ.pvdD1.Text = 0
        OZ.pvDisch1.Text = 0: OZ.pvdDov1.Text = 0
End If
vrv = 0
 ' 2B
104111: ras = 0: hop2 = BP.ph2: Xop2 = BP.pX2: Yop2 = BP.pY2: OH2 = BP.pOH2: N = 0: dNtus = 0: stre = OZ.pStre2
2151: dhh2 = (h - 750) + ((hmet - hop2) / 10)
        If zo11 = 1 Then
         xc = Xc2: yc = Yc2: hc = hc
         Else
         xc = xc: yc = yc: hc = hc
         End If
         Xc2 = xc: Yc2 = yc: hc2 = hc
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
     
snar = OZ.pSnar2: vzriv = OZ.pVzr2

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
       
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
       If snar = "ОФ" Or snar = "3ОФ56" And vzriv = "АР-5" Then dNtus2 = 0
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
      Xc2 = xc: Yc2 = yc: hc2 = hc
If BP.pX2 <> 0 Then
              OZ.pvSnar2.Text = snar: OZ.pvvzr2.Text = vzriv: OZ.pvZar2.Text = zar: OZ.pvPric2.Text = preps2: OZ.pvN2.Text = CInt(N2): OZ.pvDov2.Text = dovisch2
         OZ.pvdXtus2.Text = dXtus2: OZ.pvdNtus2.Text = dNtus2: OZ.pvPolet2.Text = ts2: OZ.pvVustra2.Text = Vustra2
        OZ.pvVd2.Text = Vd: OZ.pvDt2.Text = Dt2: OZ.pvYgt2.Text = Ygolt2: OZ.pvDovt2.Text = Dovort2: OZ.pvYr2.Text = Yr2: OZ.pvOH2.Text = OH2: OZ.pvdD2.Text = CInt(popvD)
        OZ.pvDisch2.Text = Int(Disch2): OZ.pvdDov2.Text = CInt(popvnap2)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "2 Батарея")
Else
              OZ.pvSnar2.Text = 0: OZ.pvvzr2.Text = 0: OZ.pvZar2.Text = 0: OZ.pvPric2.Text = 0: OZ.pvN2.Text = 0: OZ.pvDov2.Text = 0
         OZ.pvdXtus2.Text = 0: OZ.pvdNtus2.Text = 0: OZ.pvPolet2.Text = 0: OZ.pvVustra2.Text = 0
        OZ.pvVd2.Text = 0: OZ.pvDt2.Text = 0: OZ.pvYgt2.Text = 0: OZ.pvDovt2.Text = 0: OZ.pvYr2.Text = 0: OZ.pvOH2.Text = 0: OZ.pvdD2.Text = 0
        OZ.pvDisch2.Text = 0: OZ.pvdDov2.Text = 0
End If
vrv = 0
  '3B
501003:
1041193: ras = 0: Xop3 = BP.pX3: Yop3 = BP.pY3: hop3 = BP.ph3: OH3 = BP.pOH3: N = 0: dNtus = 0: stre = OZ.pStre3
2152: dhh3 = (h - 750) + ((hmet - hop3) / 10)
        If zo11 = 1 Then
          xc = Xc3: yc = Yc3: hc = hc
          Else
          xc = xc: yc = yc: hc = hc
          End If
          Xc3 = xc: Yc3 = yc: hc3 = hc
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
       
snar = OZ.pSnar3: vzriv = OZ.pVzr3

       If stre = "Мортирная" Then
       OZ.podPOPRMORT zar, Dt, dXtus, ybyl, ts, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, met, ybylc, Wx, Wz, dddt
       Else
       OZ.podPOPRAVKI zar, snar, Dt, Ygvozv, Ygpad, Vustra, dXtus, zc, dZwc, dXwc, dXhc, dXtc, dXv0c, Wx, Wz, dddt, Ygolt, ts, Vd
       End If
              
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
       Xc3 = xc: Yc3 = yc: hc3 = hc
If BP.pX3 <> 0 Then
                     OZ.pvSnar3.Text = snar: OZ.pvvzr3.Text = vzriv: OZ.pvZar3.Text = zar: OZ.pvPric3.Text = preps3: OZ.pvN3.Text = CInt(N3): OZ.pvDov3.Text = dovisch3
         OZ.pvdXtus3.Text = dXtus3: OZ.pvdNtus3.Text = dNtus3: OZ.pvPolet3.Text = ts3: OZ.pvVustra3.Text = Vustra3
        OZ.pvVd3.Text = Vd: OZ.pvDt3.Text = Dt3: OZ.pvYgt3.Text = Ygolt3: OZ.pvDovt3.Text = Dovort3: OZ.pvYr3.Text = Yr3: OZ.pvOH3.Text = OH3: OZ.pvdD3.Text = CInt(popvD)
        OZ.pvDisch3.Text = Int(Disch3): OZ.pvdDov3.Text = CInt(popvnap3)
If vrv = 15 Then soob = MsgBox("ВРВ больше 15", vbOKOnly, "3 Батарея")
Else
                     OZ.pvSnar3.Text = 0: OZ.pvvzr3.Text = 0: OZ.pvZar3.Text = 0: OZ.pvPric3.Text = 0: OZ.pvN3.Text = 0: OZ.pvDov3.Text = 0
         OZ.pvdXtus3.Text = 0: OZ.pvdNtus3.Text = 0: OZ.pvPolet3.Text = 0: OZ.pvVustra3.Text = 0
        OZ.pvVd3.Text = 0: OZ.pvDt3.Text = 0: OZ.pvYgt3.Text = 0: OZ.pvDovt3.Text = 0: OZ.pvYr3.Text = 0: OZ.pvOH3.Text = 0: OZ.pvdD3.Text = 0
        OZ.pvDisch3.Text = 0: OZ.pvdDov3.Text = 0
End If
vrv = 0
If Ygolzo - 1500 <= 0 Then
        Aatzo = Ygolzo + 6000 - 1500
        Else
            Aatzo = Ygolzo - 1500
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
pv1BFr.Text = Round(batych1): pv2BFr.Text = Round(batych2): pv3BFr.Text = Round(batych3)
 intve1 = Round(batych1 / (Dt1 * 0.001 + 0.001) * 0.95)
 intve2 = Round(batych2 / (Dt2 * 0.001 + 0.001) * 0.95)
 intve3 = Round(batych3 / (Dt3 * 0.001 + 0.001) * 0.95)
  Sk1 = Round(batych1 / 4 / (dXtus11 + 0.001)): Sk2 = Round(batych2 / 4 / (dXtus2 + 0.001)): Sk3 = Round(batych3 / 4 / (dXtus3 + 0.001))
  If Ygs1 > 750 Then pvZO1.Text = "ФЛАНГОВЫЙ"
       If Ygs1 <= 750 Then pvZO1.Text = "ФРОНТАЛЬНЫЙ"
       If Ygs2 > 750 Then pvZO2.Text = "ФЛАНГОВЫЙ"
       If Ygs2 <= 750 Then pvZO2.Text = "ФРОНТАЛЬНЫЙ"
       If Ygs3 > 750 Then pvZO3.Text = "ФЛАНГОВЫЙ"
       If Ygs3 <= 750 Then pvZO3.Text = "ФРОНТАЛЬНЫЙ"
       If p1bat = False Then pvZO1.Text = 0
If p2bat = False Then pvZO2.Text = 0
If p3bat = False Then pvZO3.Text = 0
       If Ygs1 > 750 Then intve1 = 0
       If Ygs1 <= 750 Then Sk1 = 0
       If Ygs2 > 750 Then intve2 = 0
       If Ygs2 <= 750 Then Sk2 = 0
       If Ygs3 > 750 Then intve3 = 0
       If Ygs3 <= 750 Then Sk3 = 0
 OZ.pvVeer1.Text = intve1: OZ.pvSk1.Text = Sk1
 OZ.pvVeer2.Text = intve2: OZ.pvSk2.Text = Sk2
 OZ.pvVeer3.Text = intve3: OZ.pvSk3.Text = Sk3
End Sub
Private Sub pvplNZO_Click()
Dim nnZo As String
Dim Xl As Single, Yl As Single, Xp As Single, Yp As Single, hc As Single
nnZo = pvplNZO
Open "D:\YO_NA\nzo" For Input As #1
1: If EOF(1) Then GoTo 10
Input #1, ta1, ta2, ta3, ta4, ta5, ta6
If ta1 = nnZo Then Xl = ta2: Yl = ta3: Xp = ta4: Yp = ta5: hc = ta6: GoTo 10
GoTo 1
10: Close #1
pXl.Text = Xl: pYl.Text = Yl: pXp.Text = Xp: pYp.Text = Yp: phc.Text = hc
End Sub
Private Sub pXp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYp.Text = ""
pYp.SetFocus
End If
End Sub
Private Sub pYp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pXl.Text = ""
pXl.SetFocus
End If
End Sub
Private Sub pXl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYl.Text = ""
pYl.SetFocus
End If
End Sub
Private Sub pYl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc.Text = ""
phc.SetFocus
End If
End Sub
Private Sub pAp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pDp.Text = ""
pDp.SetFocus
End If
End Sub
Private Sub pDp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMcp.Text = ""
pMcp.SetFocus
End If
End Sub
Private Sub pAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pDl.Text = ""
pDl.SetFocus
End If
End Sub
Private Sub pDl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pMcl.Text = ""
pMcl.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim t(1 To 10) As String
Dim i As Integer

941 Open "D:\YO_NA\nzo" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 942
 Input #1, t(1), t(2), t(3), t(4), t(5), t(6)
pvplNZO.AddItem t(1)
Loop
942 Close #1
End Sub
