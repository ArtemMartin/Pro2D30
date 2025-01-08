VERSION 5.00
Begin VB.Form Zelkagdomy6 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Цель для каждого"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frOr9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Кор3"
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
      Height          =   3200
      Left            =   7900
      TabIndex        =   74
      Top             =   6700
      Width           =   3700
      Begin VB.TextBox pXc9 
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
         Left            =   800
         TabIndex        =   78
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.TextBox pYc9 
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
         Left            =   800
         TabIndex        =   77
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox phc9 
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
         Left            =   800
         TabIndex        =   76
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.ComboBox pNc9 
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
         Left            =   2000
         TabIndex        =   75
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label Label36 
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
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   100
         TabIndex        =   82
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label35 
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
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   100
         TabIndex        =   81
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label34 
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
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   100
         TabIndex        =   80
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
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
         Height          =   400
         Left            =   100
         TabIndex        =   79
         Top             =   2400
         Width           =   1800
      End
   End
   Begin VB.Frame frOr3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Дес3"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   3200
      Left            =   100
      TabIndex        =   65
      Top             =   6700
      Width           =   3700
      Begin VB.TextBox pXc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   450
         Left            =   800
         TabIndex        =   69
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.TextBox pYc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   450
         Left            =   800
         TabIndex        =   68
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox phc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   450
         Left            =   800
         TabIndex        =   67
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.ComboBox pNc3 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   450
         Left            =   2000
         TabIndex        =   66
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label Label32 
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
         ForeColor       =   &H00004080&
         Height          =   400
         Left            =   100
         TabIndex        =   73
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label31 
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
         ForeColor       =   &H00004080&
         Height          =   400
         Left            =   100
         TabIndex        =   72
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label30 
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
         ForeColor       =   &H00004080&
         Height          =   400
         Left            =   100
         TabIndex        =   71
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   400
         Left            =   100
         TabIndex        =   70
         Top             =   2400
         Width           =   1800
      End
   End
   Begin VB.Frame frOr6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Сам3"
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
      Height          =   3200
      Left            =   4000
      TabIndex        =   56
      Top             =   6700
      Width           =   3700
      Begin VB.TextBox pXc6 
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
         Left            =   800
         TabIndex        =   60
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.TextBox pYc6 
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
         Left            =   800
         TabIndex        =   59
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox phc6 
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
         Left            =   800
         TabIndex        =   58
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.ComboBox pNc6 
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
         Left            =   2000
         TabIndex        =   57
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label Label28 
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
         ForeColor       =   &H00808000&
         Height          =   400
         Left            =   100
         TabIndex        =   64
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У"
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
         Height          =   400
         Left            =   100
         TabIndex        =   63
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label26 
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
         ForeColor       =   &H00808000&
         Height          =   400
         Left            =   100
         TabIndex        =   62
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
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
         Height          =   400
         Left            =   100
         TabIndex        =   61
         Top             =   2400
         Width           =   1800
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   1300
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8550
      Width           =   1500
   End
   Begin VB.CommandButton ZelKagdResh 
      BackColor       =   &H00FF8080&
      Caption         =   "Решить"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   100
      Width           =   2000
   End
   Begin VB.Frame frOr8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Кор2"
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
      Height          =   3200
      Left            =   7900
      TabIndex        =   40
      Top             =   3400
      Width           =   3700
      Begin VB.ComboBox pNc8 
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
         Left            =   2000
         TabIndex        =   55
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox phc8 
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
         Left            =   800
         TabIndex        =   47
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pYc8 
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
         Left            =   800
         TabIndex        =   46
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pXc8 
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
         Left            =   800
         TabIndex        =   45
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
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
         Height          =   400
         Left            =   100
         TabIndex        =   44
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label Label23 
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
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   100
         TabIndex        =   43
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label22 
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
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   100
         TabIndex        =   42
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label21 
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
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   100
         TabIndex        =   41
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame frOr7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Кор1"
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
      Height          =   3200
      Left            =   7900
      TabIndex        =   32
      Top             =   100
      Width           =   3700
      Begin VB.ComboBox pNc7 
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
         Left            =   2000
         TabIndex        =   54
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox phc7 
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
         Left            =   800
         TabIndex        =   39
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pYc7 
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
         Left            =   800
         TabIndex        =   38
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pXc7 
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
         Left            =   800
         TabIndex        =   37
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
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
         Height          =   400
         Left            =   100
         TabIndex        =   36
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label Label19 
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
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   100
         TabIndex        =   35
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label18 
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
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   100
         TabIndex        =   34
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label17 
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
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   100
         TabIndex        =   33
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame frOr5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Сам2"
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
      Height          =   3200
      Left            =   4000
      TabIndex        =   24
      Top             =   3400
      Width           =   3700
      Begin VB.ComboBox pNc5 
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
         Left            =   2000
         TabIndex        =   53
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox phc5 
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
         Left            =   800
         TabIndex        =   31
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pYc5 
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
         Left            =   800
         TabIndex        =   30
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pXc5 
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
         Left            =   800
         TabIndex        =   29
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
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
         Height          =   400
         Left            =   100
         TabIndex        =   28
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label Label15 
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
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   100
         TabIndex        =   27
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У"
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
         Height          =   400
         Left            =   100
         TabIndex        =   26
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label13 
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
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   100
         TabIndex        =   25
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame frOr4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Сам1"
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
      Height          =   3200
      Left            =   4000
      TabIndex        =   16
      Top             =   100
      Width           =   3700
      Begin VB.ComboBox pNc4 
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
         Left            =   2000
         TabIndex        =   52
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox phc4 
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
         Left            =   800
         TabIndex        =   23
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pYc4 
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
         Left            =   800
         TabIndex        =   22
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pXc4 
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
         Left            =   800
         TabIndex        =   21
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
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
         Height          =   400
         Left            =   100
         TabIndex        =   20
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H000040C0&
         Height          =   400
         Left            =   100
         TabIndex        =   19
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label10 
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
         ForeColor       =   &H000040C0&
         Height          =   400
         Left            =   100
         TabIndex        =   18
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label9 
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
         ForeColor       =   &H000040C0&
         Height          =   400
         Left            =   100
         TabIndex        =   17
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame frOr2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Дес2"
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
      Height          =   3200
      Left            =   100
      TabIndex        =   8
      Top             =   3400
      Width           =   3700
      Begin VB.ComboBox pNc2 
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
         Left            =   2000
         TabIndex        =   51
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox phc2 
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
         Left            =   800
         TabIndex        =   15
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pYc2 
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
         Left            =   800
         TabIndex        =   14
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pXc2 
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
         Left            =   800
         TabIndex        =   13
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
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
         Height          =   400
         Left            =   100
         TabIndex        =   12
         Top             =   2400
         Width           =   1800
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
         ForeColor       =   &H000000FF&
         Height          =   400
         Left            =   100
         TabIndex        =   11
         Top             =   1600
         Width           =   600
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
         ForeColor       =   &H000000FF&
         Height          =   400
         Left            =   100
         TabIndex        =   10
         Top             =   1000
         Width           =   600
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
         ForeColor       =   &H000000FF&
         Height          =   400
         Left            =   100
         TabIndex        =   9
         Top             =   400
         Width           =   600
      End
   End
   Begin VB.Frame frOr1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Дес1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3200
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   3700
      Begin VB.ComboBox pNc1 
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
         Left            =   2000
         TabIndex        =   50
         Text            =   "0"
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox phc1 
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
         Left            =   800
         TabIndex        =   7
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pYc1 
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
         Left            =   800
         TabIndex        =   6
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pXc1 
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
         Left            =   800
         TabIndex        =   5
         Text            =   "0"
         Top             =   400
         Width           =   1500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Пл. Цели"
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
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label2 
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
         TabIndex        =   2
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label1 
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
         TabIndex        =   1
         Top             =   400
         Width           =   600
      End
   End
End
Attribute VB_Name = "Zelkagdomy6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Zelkagdomy6.Hide
End Sub

Private Sub Form_Load()
Dim t(1 To 10) As String
941 Open "D:\YO_NA\zeli" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 942
 Input #1, t(1), t(2), t(3), t(4), t(5), t(6)
pNc1.AddItem t(1)
pNc2.AddItem t(1)
pNc3.AddItem t(1)
pNc4.AddItem t(1)
pNc5.AddItem t(1)
pNc6.AddItem t(1)
pNc7.AddItem t(1)
pNc8.AddItem t(1)
pNc9.AddItem t(1)
Loop
942 Close #1

frOr1.Caption = Shest6Oryd.labOr1
frOr2.Caption = Shest6Oryd.labOr2
frOr3.Caption = Shest6Oryd.labOr3
frOr4.Caption = Shest6Oryd.labOr4
frOr5.Caption = Shest6Oryd.labOr5
frOr6.Caption = Shest6Oryd.labOr6
frOr7.Caption = Shest6Oryd.labOr7
frOr8.Caption = Shest6Oryd.labOr8
frOr9.Caption = Shest6Oryd.labOr9

End Sub

Private Sub pNc1_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc1
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc1.Text = Xc: pYc1.Text = Yc: phc1.Text = hc
End Sub
Private Sub pNc1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc1, Xc, Yc, hc

pXc1.Text = Xc: pYc1.Text = Yc: phc1.Text = hc

End Sub
Private Sub pNc2_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc2
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc2.Text = Xc: pYc2.Text = Yc: phc2.Text = hc
End Sub
Private Sub pNc2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc2, Xc, Yc, hc

pXc2.Text = Xc: pYc2.Text = Yc: phc2.Text = hc

End Sub

Private Sub pNc3_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc3
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc3.Text = Xc: pYc3.Text = Yc: phc3.Text = hc
End Sub
Private Sub pNc3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc3, Xc, Yc, hc

pXc3.Text = Xc: pYc3.Text = Yc: phc3.Text = hc

End Sub


Private Sub pNc4_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc4
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc4.Text = Xc: pYc4.Text = Yc: phc4.Text = hc
End Sub
Private Sub pNc4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc4, Xc, Yc, hc

pXc4.Text = Xc: pYc4.Text = Yc: phc4.Text = hc

End Sub

Private Sub pNc5_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc5
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc5.Text = Xc: pYc5.Text = Yc: phc5.Text = hc
End Sub
Private Sub pNc5_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc5, Xc, Yc, hc

pXc5.Text = Xc: pYc5.Text = Yc: phc5.Text = hc

End Sub

Private Sub pNc6_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc6
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc6.Text = Xc: pYc6.Text = Yc: phc6.Text = hc
End Sub
Private Sub pNc6_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc6, Xc, Yc, hc

pXc6.Text = Xc: pYc6.Text = Yc: phc6.Text = hc

End Sub

Private Sub pNc7_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc7
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc7.Text = Xc: pYc7.Text = Yc: phc7.Text = hc
End Sub
Private Sub pNc7_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc7, Xc, Yc, hc

pXc7.Text = Xc: pYc7.Text = Yc: phc7.Text = hc

End Sub

Private Sub pNc8_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc8
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc8.Text = Xc: pYc8.Text = Yc: phc8.Text = hc
End Sub
Private Sub pNc8_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc8, Xc, Yc, hc

pXc8.Text = Xc: pYc8.Text = Yc: phc8.Text = hc

End Sub

Private Sub pNc9_Click()
Dim nz As String
Dim t(10) As String
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single
nz = pNc9
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = Val(t(3)): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
pXc9.Text = Xc: pYc9.Text = Yc: phc9.Text = hc
End Sub
Private Sub pNc9_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Xc As Single, Yc As Single, hc As Single, Frc As Single, Glc As Single

activityKeyDown KeyCode, pNc9, Xc, Yc, hc

pXc9.Text = Xc: pYc9.Text = Yc: phc9.Text = hc

End Sub

Private Sub pXc1_Click()
pXc1.Text = ""
End Sub

Private Sub pXc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc1.Text = ""
pYc1.SetFocus
End If
End Sub

Private Sub pXc2_Click()
pXc2.Text = ""
End Sub

Private Sub pXc3_Click()
pXc3.Text = ""
End Sub

Private Sub pXc4_Click()
pXc4.Text = ""
End Sub

Private Sub pXc5_Click()
pXc5.Text = ""
End Sub

Private Sub pYc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc1.Text = ""
phc1.SetFocus
End If
End Sub
Private Sub pXc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc2.Text = ""
pYc2.SetFocus
End If
End Sub
Private Sub pYc2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc2.Text = ""
phc2.SetFocus
End If
End Sub
Private Sub pXc3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc3.Text = ""
pYc3.SetFocus
End If
End Sub
Private Sub pYc3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc3.Text = ""
phc3.SetFocus
End If
End Sub
Private Sub pXc4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc4.Text = ""
pYc4.SetFocus
End If
End Sub
Private Sub pYc4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc4.Text = ""
phc4.SetFocus
End If
End Sub
Private Sub pXc5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc5.Text = ""
pYc5.SetFocus
End If
End Sub
Private Sub pYc5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc5.Text = ""
phc5.SetFocus
End If
End Sub
Private Sub pXc6_Click()
pXc6.Text = ""
End Sub
Private Sub pXc6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc6.Text = ""
pYc6.SetFocus
End If
End Sub
Private Sub pYc6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc6.Text = ""
phc6.SetFocus
End If
End Sub
Private Sub pXc7_Click()
pXc7.Text = ""
End Sub
Private Sub pXc7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc7.Text = ""
pYc7.SetFocus
End If
End Sub
Private Sub pYc7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc7.Text = ""
phc7.SetFocus
End If
End Sub
Private Sub pXc8_Click()
pXc8.Text = ""
End Sub
Private Sub pXc8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc8.Text = ""
pYc8.SetFocus
End If
End Sub
Private Sub pYc8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc8.Text = ""
phc8.SetFocus
End If
End Sub
Private Sub pXc9_Click()
pXc9.Text = ""
End Sub
Private Sub pXc9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYc9.Text = ""
pYc9.SetFocus
End If
End Sub
Private Sub pYc9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc9.Text = ""
phc9.SetFocus
End If
End Sub

Private Sub ZelKagdResh_Click()
 Xc = pXc1: Yc = pYc1: hc = phc1
Dim Nop As Integer
Dim Xop As Single
Xop = bp6Oryd.pAks1X: Nop = 1: rep = Shest6Oryd.pRep1
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric1.Text = preps1: Shest6Oryd.pvN1.Text = Round(N1): Shest6Oryd.pvDov1.Text = dovisch1
    Shest6Oryd.pvdXtus1.Text = dXtus: Shest6Oryd.pvts1.Text = ts: Shest6Oryd.pvDt1.Text = Dt1: Shest6Oryd.pvYgt1.Text = Ygolt
    Shest6Oryd.pvDovt1.Text = Dovort1: Shest6Oryd.pvYr1.Text = Round(yroven): Shest6Oryd.pvdD1.Text = Round(popvd1)
    Shest6Oryd.pvdDov1.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric1.Text = 0: Shest6Oryd.pvN1.Text = 0: Shest6Oryd.pvDov1.Text = 0: Shest6Oryd.pvdXtus1.Text = 0: Shest6Oryd.pvts1.Text = 0
    Shest6Oryd.pvDt1.Text = 0: Shest6Oryd.pvYgt1.Text = 0: Shest6Oryd.pvDovt1.Text = 0: Shest6Oryd.pvYr1.Text = 0: Shest6Oryd.pvdD1.Text = 0
    Shest6Oryd.pvdDov1.Text = 0
End If
Xc = pXc2: Yc = pYc2: hc = phc2
Xop = bp6Oryd.pAks2X: Nop = 2: rep = Shest6Oryd.pRep2
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric2.Text = preps1: Shest6Oryd.pvN2.Text = Round(N1): Shest6Oryd.pvDov2.Text = dovisch1
    Shest6Oryd.pvdXtus2.Text = dXtus: Shest6Oryd.pvts2.Text = ts: Shest6Oryd.pvDt2.Text = Dt1: Shest6Oryd.pvYgt2.Text = Ygolt
    Shest6Oryd.pvDovt2.Text = Dovort1: Shest6Oryd.pvYr2.Text = Round(yroven): Shest6Oryd.pvdD2.Text = Round(popvd1)
    Shest6Oryd.pvdDov2.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric2.Text = 0: Shest6Oryd.pvN2.Text = 0: Shest6Oryd.pvDov2.Text = 0: Shest6Oryd.pvdXtus2.Text = 0: Shest6Oryd.pvts2.Text = 0
    Shest6Oryd.pvDt2.Text = 0: Shest6Oryd.pvYgt2.Text = 0: Shest6Oryd.pvDovt2.Text = 0: Shest6Oryd.pvYr2.Text = 0: Shest6Oryd.pvdD2.Text = 0
    Shest6Oryd.pvdDov2.Text = 0
End If
Xc = pXc3: Yc = pYc3: hc = phc3
Xop = bp6Oryd.pAks3X: Nop = 3: rep = Shest6Oryd.pRep3
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric3.Text = preps1: Shest6Oryd.pvN3.Text = Round(N1): Shest6Oryd.pvDov3.Text = dovisch1
    Shest6Oryd.pvdXtus3.Text = dXtus: Shest6Oryd.pvts3.Text = ts: Shest6Oryd.pvDt3.Text = Dt1: Shest6Oryd.pvYgt3.Text = Ygolt
    Shest6Oryd.pvDovt3.Text = Dovort1: Shest6Oryd.pvYr3.Text = Round(yroven): Shest6Oryd.pvdD3.Text = Round(popvd1)
    Shest6Oryd.pvdDov3.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric3.Text = 0: Shest6Oryd.pvN3.Text = 0: Shest6Oryd.pvDov3.Text = 0: Shest6Oryd.pvdXtus3.Text = 0: Shest6Oryd.pvts3.Text = 0
    Shest6Oryd.pvDt3.Text = 0: Shest6Oryd.pvYgt3.Text = 0: Shest6Oryd.pvDovt3.Text = 0: Shest6Oryd.pvYr3.Text = 0: Shest6Oryd.pvdD3.Text = 0
    Shest6Oryd.pvdDov3.Text = 0
End If
Xc = pXc4: Yc = pYc4: hc = phc4
Xop = bp6Oryd.pKal1X: Nop = 4: rep = Shest6Oryd.pRep4
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric4.Text = preps1: Shest6Oryd.pvN4.Text = Round(N1): Shest6Oryd.pvDov4.Text = dovisch1
    Shest6Oryd.pvdXtus4.Text = dXtus: Shest6Oryd.pvts4.Text = ts: Shest6Oryd.pvDt4.Text = Dt1: Shest6Oryd.pvYgt4.Text = Ygolt
    Shest6Oryd.pvDovt4.Text = Dovort1: Shest6Oryd.pvYr4.Text = Round(yroven): Shest6Oryd.pvdD4.Text = Round(popvd1)
    Shest6Oryd.pvdDov4.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric4.Text = 0: Shest6Oryd.pvN4.Text = 0: Shest6Oryd.pvDov4.Text = 0: Shest6Oryd.pvdXtus4.Text = 0: Shest6Oryd.pvts4.Text = 0
    Shest6Oryd.pvDt4.Text = 0: Shest6Oryd.pvYgt4.Text = 0: Shest6Oryd.pvDovt4.Text = 0: Shest6Oryd.pvYr4.Text = 0: Shest6Oryd.pvdD4.Text = 0
    Shest6Oryd.pvdDov4.Text = 0
End If
Xc = pXc5: Yc = pYc5: hc = phc5
Xop = bp6Oryd.pKal2X: Nop = 5: rep = Shest6Oryd.pRep5
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric5.Text = preps1: Shest6Oryd.pvN5.Text = Round(N1): Shest6Oryd.pvDov5.Text = dovisch1
    Shest6Oryd.pvdXtus5.Text = dXtus: Shest6Oryd.pvts5.Text = ts: Shest6Oryd.pvDt5.Text = Dt1: Shest6Oryd.pvYgt5.Text = Ygolt
    Shest6Oryd.pvDovt5.Text = Dovort1: Shest6Oryd.pvYr5.Text = Round(yroven): Shest6Oryd.pvdD5.Text = Round(popvd1)
    Shest6Oryd.pvdDov5.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric5.Text = 0: Shest6Oryd.pvN5.Text = 0: Shest6Oryd.pvDov5.Text = 0: Shest6Oryd.pvdXtus5.Text = 0: Shest6Oryd.pvts5.Text = 0
    Shest6Oryd.pvDt5.Text = 0: Shest6Oryd.pvYgt5.Text = 0: Shest6Oryd.pvDovt5.Text = 0: Shest6Oryd.pvYr5.Text = 0: Shest6Oryd.pvdD5.Text = 0
    Shest6Oryd.pvdDov5.Text = 0
End If
Xc = pXc6: Yc = pYc6: hc = phc6
Xop = bp6Oryd.pKal3X: Nop = 6: rep = Shest6Oryd.pRep6
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric6.Text = preps1: Shest6Oryd.pvN6.Text = Round(N1): Shest6Oryd.pvDov6.Text = dovisch1
    Shest6Oryd.pvdXtus6.Text = dXtus: Shest6Oryd.pvts6.Text = ts: Shest6Oryd.pvDt6.Text = Dt1: Shest6Oryd.pvYgt6.Text = Ygolt
    Shest6Oryd.pvDovt6.Text = Dovort1: Shest6Oryd.pvYr6.Text = Round(yroven): Shest6Oryd.pvdD6.Text = Round(popvd1)
    Shest6Oryd.pvdDov6.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric6.Text = 0: Shest6Oryd.pvN6.Text = 0: Shest6Oryd.pvDov6.Text = 0: Shest6Oryd.pvdXtus6.Text = 0: Shest6Oryd.pvts6.Text = 0
    Shest6Oryd.pvDt6.Text = 0: Shest6Oryd.pvYgt6.Text = 0: Shest6Oryd.pvDovt6.Text = 0: Shest6Oryd.pvYr6.Text = 0: Shest6Oryd.pvdD6.Text = 0
    Shest6Oryd.pvdDov6.Text = 0
End If
Xc = pXc7: Yc = pYc7: hc = phc7
Xop = bp6Oryd.pOsk1X: Nop = 7: rep = Shest6Oryd.pRep7
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric7.Text = preps1: Shest6Oryd.pvN7.Text = Round(N1): Shest6Oryd.pvDov7.Text = dovisch1
    Shest6Oryd.pvdXtus7.Text = dXtus: Shest6Oryd.pvts7.Text = ts: Shest6Oryd.pvDt7.Text = Dt1: Shest6Oryd.pvYgt7.Text = Ygolt
    Shest6Oryd.pvDovt7.Text = Dovort1: Shest6Oryd.pvYr7.Text = Round(yroven): Shest6Oryd.pvdD7.Text = Round(popvd1)
    Shest6Oryd.pvdDov7.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric7.Text = 0: Shest6Oryd.pvN7.Text = 0: Shest6Oryd.pvDov7.Text = 0: Shest6Oryd.pvdXtus7.Text = 0: Shest6Oryd.pvts7.Text = 0
    Shest6Oryd.pvDt7.Text = 0: Shest6Oryd.pvYgt7.Text = 0: Shest6Oryd.pvDovt7.Text = 0: Shest6Oryd.pvYr7.Text = 0: Shest6Oryd.pvdD7.Text = 0
    Shest6Oryd.pvdDov7.Text = 0
End If
Xc = pXc8: Yc = pYc8: hc = phc8
Xop = bp6Oryd.pOsk2X: Nop = 8: rep = Shest6Oryd.pRep8
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric8.Text = preps1: Shest6Oryd.pvN8.Text = Round(N1): Shest6Oryd.pvDov8.Text = dovisch1
    Shest6Oryd.pvdXtus8.Text = dXtus: Shest6Oryd.pvts8.Text = ts: Shest6Oryd.pvDt8.Text = Dt1: Shest6Oryd.pvYgt8.Text = Ygolt
    Shest6Oryd.pvDovt8.Text = Dovort1: Shest6Oryd.pvYr8.Text = Round(yroven): Shest6Oryd.pvdD8.Text = Round(popvd1)
    Shest6Oryd.pvdDov8.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric8.Text = 0: Shest6Oryd.pvN8.Text = 0: Shest6Oryd.pvDov8.Text = 0: Shest6Oryd.pvdXtus8.Text = 0: Shest6Oryd.pvts8.Text = 0
    Shest6Oryd.pvDt8.Text = 0: Shest6Oryd.pvYgt8.Text = 0: Shest6Oryd.pvDovt8.Text = 0: Shest6Oryd.pvYr8.Text = 0: Shest6Oryd.pvdD8.Text = 0
    Shest6Oryd.pvdDov8.Text = 0
End If
Xc = pXc9: Yc = pYc9: hc = phc9
Xop = bp6Oryd.pOsk3X: Nop = 9: rep = Shest6Oryd.pRep9
If Xop <> 0 And Xc <> 0 Then
    Shest6Oryd.OZ6or Nop, Xc, Yc, hc, rep, preps1, N1, dovisch1, ts, dXtus, Dt1, Ygolt, Dovort1, yroven, popvd1, popvnap1
    Shest6Oryd.pvPric9.Text = preps1: Shest6Oryd.pvN9.Text = Round(N1): Shest6Oryd.pvDov9.Text = dovisch1
    Shest6Oryd.pvdXtus9.Text = dXtus: Shest6Oryd.pvts9.Text = ts: Shest6Oryd.pvDt9.Text = Dt1: Shest6Oryd.pvYgt9.Text = Ygolt
    Shest6Oryd.pvDovt9.Text = Dovort1: Shest6Oryd.pvYr9.Text = Round(yroven): Shest6Oryd.pvdD9.Text = Round(popvd1)
    Shest6Oryd.pvdDov9.Text = Round(popvnap1)
Else
    Shest6Oryd.pvPric9.Text = 0: Shest6Oryd.pvN9.Text = 0: Shest6Oryd.pvDov9.Text = 0: Shest6Oryd.pvdXtus9.Text = 0: Shest6Oryd.pvts9.Text = 0
    Shest6Oryd.pvDt9.Text = 0: Shest6Oryd.pvYgt9.Text = 0: Shest6Oryd.pvDovt9.Text = 0: Shest6Oryd.pvYr9.Text = 0: Shest6Oryd.pvdD9.Text = 0
    Shest6Oryd.pvdDov9.Text = 0
End If

'записываем номера целей
Open App.Path & "\writeZeliEach" For Output As #1
Write #1, pNc1, pNc2, pNc3, pNc4, pNc5, pNc6, pNc7, pNc8, pNc9
Close #1

End Sub
Sub activityKeyDown(ByVal kode As Integer, ByVal nz As String, Xc, Yc, hc)
Dim t(0 To 10) As String

If kode = 13 Then
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, t(0), t(1), t(2), t(3), t(4), t(5)
   If t(0) = nz Then
        Xc = t(1): Yc = t(2): hc = t(3): Fr = t(4): Gl = t(5)
        Else
            GoTo 101111
        End If
1012 Close #1
    Else
End If

End Sub

