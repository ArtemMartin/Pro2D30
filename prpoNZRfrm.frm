VERSION 5.00
Begin VB.Form prpoNZRfrm 
   Caption         =   "Пристрелка по НЗР"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16845
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   16845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   100
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6700
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
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
      Left            =   100
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5000
      Width           =   2000
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
      Height          =   2500
      Left            =   100
      TabIndex        =   51
      Top             =   2300
      Width           =   2000
      Begin VB.OptionButton pkagdomy 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   700
         TabIndex        =   55
         Top             =   1800
         Width           =   495
      End
      Begin VB.OptionButton pvsem 
         BackColor       =   &H00C0C0C0&
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
         Left            =   700
         TabIndex        =   53
         Top             =   800
         Value           =   -1  'True
         Width           =   400
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Каждому"
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
         Left            =   300
         TabIndex        =   54
         Top             =   1400
         Width           =   1300
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Всем"
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
         Left            =   500
         TabIndex        =   52
         Top             =   400
         Width           =   800
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Корректура"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2300
      TabIndex        =   23
      Top             =   3000
      Width           =   14400
      Begin VB.TextBox pvdN8 
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
         Left            =   11200
         TabIndex        =   88
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvdN7 
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
         Left            =   9800
         TabIndex        =   87
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov9 
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
         Left            =   12600
         TabIndex        =   86
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdN9 
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
         Left            =   12600
         TabIndex        =   85
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr9 
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
         Left            =   12600
         TabIndex        =   84
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD9 
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
         Height          =   450
         Left            =   12600
         TabIndex        =   83
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvdN5 
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
         Left            =   7000
         TabIndex        =   81
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvdN4 
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
         Left            =   5600
         TabIndex        =   80
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov6 
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
         Left            =   8400
         TabIndex        =   79
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdN6 
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
         Left            =   8400
         TabIndex        =   78
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr6 
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
         Left            =   8400
         TabIndex        =   77
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD6 
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
         Height          =   450
         Left            =   8400
         TabIndex        =   76
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvdN2 
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
         Left            =   2800
         TabIndex        =   73
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvdN1 
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
         Left            =   1400
         TabIndex        =   72
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov3 
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
         Left            =   4200
         TabIndex        =   71
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdN3 
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
         Left            =   4200
         TabIndex        =   70
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr3 
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
         Left            =   4200
         TabIndex        =   69
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD3 
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
         Height          =   450
         Left            =   4200
         TabIndex        =   68
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov8 
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
         Left            =   11200
         TabIndex        =   50
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov7 
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
         Left            =   9800
         TabIndex        =   49
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov5 
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
         Left            =   7000
         TabIndex        =   48
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov4 
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
         Left            =   5600
         TabIndex        =   47
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov2 
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
         Left            =   2800
         TabIndex        =   46
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov1 
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
         Left            =   1400
         TabIndex        =   45
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr8 
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
         Left            =   11200
         TabIndex        =   44
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD8 
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
         Height          =   450
         Left            =   11200
         TabIndex        =   43
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr7 
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
         Left            =   9800
         TabIndex        =   42
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD7 
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
         Height          =   450
         Left            =   9800
         TabIndex        =   41
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr5 
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
         Left            =   7000
         TabIndex        =   40
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD5 
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
         Height          =   450
         Left            =   7000
         TabIndex        =   39
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr4 
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
         Left            =   5600
         TabIndex        =   38
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD4 
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
         Height          =   450
         Left            =   5600
         TabIndex        =   37
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr2 
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
         Left            =   2800
         TabIndex        =   36
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD2 
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
         Height          =   450
         Left            =   2800
         TabIndex        =   35
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr1 
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
         Left            =   1400
         TabIndex        =   34
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvkorD1 
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
         Left            =   1400
         TabIndex        =   33
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.Label labeOr9 
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
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   12600
         TabIndex        =   82
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labeOr6 
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
         Height          =   405
         Left            =   8400
         TabIndex        =   75
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dN="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   74
         Top             =   2200
         Width           =   800
      End
      Begin VB.Label labeOr3 
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
         ForeColor       =   &H00404080&
         Height          =   400
         Left            =   4200
         TabIndex        =   67
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dДов="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   32
         Top             =   2800
         Width           =   1100
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dПр="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   31
         Top             =   1600
         Width           =   800
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД="
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
         TabIndex        =   30
         Top             =   1000
         Width           =   800
      End
      Begin VB.Label labeOr8 
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
         ForeColor       =   &H00800000&
         Height          =   400
         Left            =   11200
         TabIndex        =   29
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labeOr7 
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
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   9800
         TabIndex        =   28
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labeOr5 
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
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   7000
         TabIndex        =   27
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labeOr4 
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
         Height          =   400
         Left            =   5600
         TabIndex        =   26
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labeOr2 
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
         Height          =   400
         Left            =   2800
         TabIndex        =   25
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labeOr1 
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
         Height          =   400
         Left            =   1400
         TabIndex        =   24
         Top             =   400
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "По разрыву"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   2300
      TabIndex        =   2
      Top             =   100
      Width           =   14400
      Begin VB.TextBox pdD9 
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
         Left            =   13000
         TabIndex        =   66
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA9 
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
         Left            =   13000
         TabIndex        =   65
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdD6 
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
         Left            =   8500
         TabIndex        =   63
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA6 
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
         Left            =   8500
         TabIndex        =   62
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdD3 
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
         Left            =   4000
         TabIndex        =   60
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA3 
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
         Left            =   4000
         TabIndex        =   59
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdD8 
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
         Height          =   450
         Left            =   11500
         TabIndex        =   22
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA8 
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
         Height          =   450
         Left            =   11500
         TabIndex        =   21
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdD7 
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
         Left            =   10000
         TabIndex        =   20
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA7 
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
         Left            =   10000
         TabIndex        =   19
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdD5 
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
         Left            =   7000
         TabIndex        =   18
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA5 
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
         Left            =   7000
         TabIndex        =   17
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdD4 
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
         Left            =   5500
         TabIndex        =   16
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA4 
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
         Left            =   5500
         TabIndex        =   15
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdD2 
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
         Left            =   2500
         TabIndex        =   14
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA2 
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
         Left            =   2500
         TabIndex        =   13
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdD1 
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
         Left            =   1000
         TabIndex        =   12
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdA1 
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
         Left            =   1000
         TabIndex        =   11
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.Label labOr9 
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
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   13000
         TabIndex        =   64
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labOr6 
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
         Height          =   400
         Left            =   8500
         TabIndex        =   61
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labOr3 
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
         ForeColor       =   &H00404080&
         Height          =   400
         Left            =   4000
         TabIndex        =   58
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД="
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
         TabIndex        =   10
         Top             =   1600
         Width           =   600
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dA="
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
         TabIndex        =   9
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label labOr8 
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
         ForeColor       =   &H00800000&
         Height          =   400
         Left            =   11500
         TabIndex        =   8
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labOr7 
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
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   10000
         TabIndex        =   7
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labOr5 
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
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   7000
         TabIndex        =   6
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labOr4 
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
         Height          =   400
         Left            =   5500
         TabIndex        =   5
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labOr2 
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
         Height          =   400
         Left            =   2500
         TabIndex        =   4
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label labOr1 
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
         Height          =   400
         Left            =   1000
         TabIndex        =   3
         Top             =   400
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "№ КНП"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   2000
      Begin VB.ComboBox pNKP 
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
         ItemData        =   "prpoNZRfrm.frx":0000
         Left            =   480
         List            =   "prpoNZRfrm.frx":0013
         TabIndex        =   1
         Text            =   "1"
         Top             =   720
         Width           =   800
      End
   End
End
Attribute VB_Name = "prpoNZRfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim Arr As Single, Dr As Single, Mr As Single, nkp As Single, nb As Single, Xkp As Single, Ykp As Single, hkp As Single, xc As Single, yc As Single, hc As Single
 Dim Ak As Single
  nkp = pNKP
   If nkp = 1 Then Xkp = BP.pXkp1: Ykp = BP.pYkp1: hkp = BP.phkp1
   If nkp = 2 Then Xkp = BP.pXkp2: Ykp = BP.pYkp2: hkp = BP.phkp2
   If nkp = 3 Then Xkp = BP.pXkp3: Ykp = BP.pYkp3: hkp = BP.phkp3
   If nkp = 4 Then Xkp = BP.pXkp4: Ykp = BP.pYkp4: hkp = BP.phkp4
   If nkp = 5 Then Xkp = BP.pXkp5: Ykp = BP.pYkp5: hkp = BP.phkp5
   
   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc1: yc = Zelkagdomy6.pYc1
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD1, pdA1, Dr, Arr
 If pdD1 = 0 And pdA1 = 0 Then
    Else
        Xb = bp6Oryd.pAks1X: Yb = bp6Oryd.pAks1Y: hb = bp6Oryd.pAks1h: dXtus = Shest6Oryd.pvdXtus1: Dt = Shest6Oryd.pvDt1
        Ygolt = Shest6Oryd.pvYgt1: snar = Shest6Oryd.pAks1Snar: zar = Shest6Oryd.pAks1Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD1 = dD: pvkorDov1 = dDov: pvkorPr1 = dPr: pvdN1 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc2: yc = Zelkagdomy6.pYc2
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD2, pdA2, Dr, Arr
  If pdD2 = 0 And pdA2 = 0 Then
    Else
        Xb = bp6Oryd.pAks2X: Yb = bp6Oryd.pAks2Y: hb = bp6Oryd.pAks2h
        dXtus = Shest6Oryd.pvdXtus2: Dt = Shest6Oryd.pvDt2: Ygolt = Shest6Oryd.pvYgt2: snar = Shest6Oryd.pAks2Snar: zar = Shest6Oryd.pAks2Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD2 = dD: pvkorDov2 = dDov: pvkorPr2 = dPr: pvdN2 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc3: yc = Zelkagdomy6.pYc3
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD3, pdA3, Dr, Arr
  If pdD3 = 0 And pdA3 = 0 Then
    Else
        Xb = bp6Oryd.pAks3X: Yb = bp6Oryd.pAks3Y: hb = bp6Oryd.pAks3h
        dXtus = Shest6Oryd.pvdXtus3: Dt = Shest6Oryd.pvDt3: Ygolt = Shest6Oryd.pvYgt3: snar = Shest6Oryd.pAks3Snar: zar = Shest6Oryd.pAks3Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD3 = dD: pvkorDov3 = dDov: pvkorPr3 = dPr: pvdN3 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc4: yc = Zelkagdomy6.pYc4
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD4, pdA4, Dr, Arr
  If pdD4 = 0 And pdA4 = 0 Then
    Else
        Xb = bp6Oryd.pKal1X: Yb = bp6Oryd.pKal1Y: hb = bp6Oryd.pKal1h
        dXtus = Shest6Oryd.pvdXtus4: Dt = Shest6Oryd.pvDt4: Ygolt = Shest6Oryd.pvYgt4: snar = Shest6Oryd.pKal1Snar: zar = Shest6Oryd.pKal1Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD4 = dD: pvkorDov4 = dDov: pvkorPr4 = dPr: pvdN4 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc5: yc = Zelkagdomy6.pYc5
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD5, pdA5, Dr, Arr
  If pdD5 = 0 And pdA5 = 0 Then
    Else
        Xb = bp6Oryd.pKal2X: Yb = bp6Oryd.pKal2Y: hb = bp6Oryd.pKal2h
        dXtus = Shest6Oryd.pvdXtus5: Dt = Shest6Oryd.pvDt5: Ygolt = Shest6Oryd.pvYgt5: snar = Shest6Oryd.pKal2Snar: zar = Shest6Oryd.pKal2Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD5 = dD: pvkorDov5 = dDov: pvkorPr5 = dPr: pvdN5 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc6: yc = Zelkagdomy6.pYc6
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD6, pdA6, Dr, Arr
  If pdD6 = 0 And pdA6 = 0 Then
    Else
        Xb = bp6Oryd.pKal3X: Yb = bp6Oryd.pKal3Y: hb = bp6Oryd.pKal3h
        dXtus = Shest6Oryd.pvdXtus6: Dt = Shest6Oryd.pvDt6: Ygolt = Shest6Oryd.pvYgt6: snar = Shest6Oryd.pKal3Snar: zar = Shest6Oryd.pKal3Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD6 = dD: pvkorDov6 = dDov: pvkorPr6 = dPr: pvdN6 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc7: yc = Zelkagdomy6.pYc7
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD7, pdA7, Dr, Arr
  If pdD7 = 0 And pdA7 = 0 Then
    Else
        Xb = bp6Oryd.pOsk1X: Yb = bp6Oryd.pOsk1Y: hb = bp6Oryd.pOsk1h
        dXtus = Shest6Oryd.pvdXtus7: Dt = Shest6Oryd.pvDt7: Ygolt = Shest6Oryd.pvYgt7: snar = Shest6Oryd.pOsk1Snar: zar = Shest6Oryd.pOsk1Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD7 = dD: pvkorDov7 = dDov: pvkorPr7 = dPr: pvdN7 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc8: yc = Zelkagdomy6.pYc8
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD8, pdA8, Dr, Arr
  If pdD8 = 0 And pdA8 = 0 Then
    Else
        Xb = bp6Oryd.pOsk2X: Yb = bp6Oryd.pOsk2Y: hb = bp6Oryd.pOsk2h
        dXtus = Shest6Oryd.pvdXtus8: Dt = Shest6Oryd.pvDt8: Ygolt = Shest6Oryd.pvYgt8: snar = Shest6Oryd.pOsk2Snar: zar = Shest6Oryd.pOsk2Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD8 = dD: pvkorDov8 = dDov: pvkorPr8 = dPr: pvdN8 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc9: yc = Zelkagdomy6.pYc9
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
DrAr xc, yc, Xkp, Ykp, pdD9, pdA9, Dr, Arr
  If pdD9 = 0 And pdA9 = 0 Then
    Else
        Xb = bp6Oryd.pOsk3X: Yb = bp6Oryd.pOsk3Y: hb = bp6Oryd.pOsk3h
        dXtus = Shest6Oryd.pvdXtus9: Dt = Shest6Oryd.pvDt9: Ygolt = Shest6Oryd.pvYgt9: snar = Shest6Oryd.pOsk3Snar: zar = Shest6Oryd.pOsk3Zar
        PrisDAKform.podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD9 = dD: pvkorDov9 = dDov: pvkorPr9 = dPr: pvdN9 = dN
End If

End Sub

Private Sub Command2_Click()
prpoNZRfrm.Hide
End Sub

Private Sub Form_Load()
labOr1 = Shest6Oryd.labOr1
labOr2 = Shest6Oryd.labOr2
labOr3 = Shest6Oryd.labOr3
labOr4 = Shest6Oryd.labOr4
labOr5 = Shest6Oryd.labOr5
labOr6 = Shest6Oryd.labOr6
labOr7 = Shest6Oryd.labOr7
labOr8 = Shest6Oryd.labOr8
labOr9 = Shest6Oryd.labOr9

labeOr1 = Shest6Oryd.labOr1
labeOr2 = Shest6Oryd.labOr2
labeOr3 = Shest6Oryd.labOr3
labeOr4 = Shest6Oryd.labOr4
labeOr5 = Shest6Oryd.labOr5
labeOr6 = Shest6Oryd.labOr6
labeOr7 = Shest6Oryd.labOr7
labeOr8 = Shest6Oryd.labOr8
labeOr9 = Shest6Oryd.labOr9
End Sub

Private Sub pdA1_Click()
pdA1.Text = ""
End Sub
Private Sub pdA2_Click()
pdA2.Text = ""
End Sub
Private Sub pdA3_Click()
pdA3.Text = ""
End Sub
Private Sub pdA4_Click()
pdA4.Text = ""
End Sub
Private Sub pdA5_Click()
pdA5.Text = ""
End Sub
Private Sub pdA6_Click()
pdA6.Text = ""
End Sub
Private Sub pdA7_Click()
pdA7.Text = ""
End Sub
Private Sub pdA8_Click()
pdA8.Text = ""
End Sub
Private Sub pdA9_Click()
pdA9.Text = ""
End Sub

Private Sub pdA1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdD1.Text = ""
pdD1.SetFocus
End If
End Sub
Private Sub pdA2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdD2.Text = ""
pdD2.SetFocus
End If
End Sub
Private Sub pdA3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdD3.Text = ""
pdD3.SetFocus
End If
End Sub
Private Sub pdA4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdD4.Text = ""
pdD4.SetFocus
End If
End Sub
Private Sub pdA5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdD5.Text = ""
pdD5.SetFocus
End If
End Sub
Private Sub pdA6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdD6.Text = ""
pdD6.SetFocus
End If
End Sub
Sub DrAr(ByVal xc As Single, ByVal yc As Single, ByVal Xkp As Single, ByVal Ykp As Single, ByVal dD As Single, ByVal dA As Single, Dr, Arr)
        dx = xc - Xkp
        dy = yc - Ykp
 Dk = Sqr(dx ^ 2 + dy ^ 2)
 Ar = Abs(Atn(dy / (dx + 0.1)) / 3.141592 * 30) * 100
 If dx > 0 And dy > 0 Then Ak = Int(Ar)
 If dx < 0 And dy > 0 Then Ak = Int(3000 - Ar)
 If dx < 0 And dy < 0 Then Ak = Int(3000 + Ar)
 If dx > 0 And dy < 0 Then Ak = Int(6000 - Ar)
 Dr = dD + Dk: Arr = dA + Ak
End Sub
