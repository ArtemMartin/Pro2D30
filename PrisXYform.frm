VERSION 5.00
Begin VB.Form PrisXYform 
   Caption         =   "Пристрелка Х, У"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   18555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   16600
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5200
      Width           =   1600
   End
   Begin VB.Frame Frame2 
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
      Height          =   3660
      Left            =   100
      TabIndex        =   22
      Top             =   2900
      Width           =   14145
      Begin VB.TextBox pvkorDov9 
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
         Left            =   12700
         TabIndex        =   81
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkordN9 
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
         Left            =   12700
         TabIndex        =   80
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr9 
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
         Left            =   12700
         TabIndex        =   79
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
         Left            =   12700
         TabIndex        =   78
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkordN8 
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
         Left            =   11300
         TabIndex        =   76
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkordN7 
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
         Left            =   9700
         TabIndex        =   75
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov6 
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
         Left            =   8300
         TabIndex        =   74
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkordN6 
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
         Left            =   8300
         TabIndex        =   73
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr6 
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
         Left            =   8300
         TabIndex        =   72
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
         Left            =   8300
         TabIndex        =   71
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkordN5 
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
         Left            =   6900
         TabIndex        =   69
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkordN4 
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
         Left            =   5500
         TabIndex        =   68
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov3 
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
         Left            =   4100
         TabIndex        =   67
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkordN3 
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
         Left            =   4100
         TabIndex        =   66
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr3 
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
         Left            =   4100
         TabIndex        =   65
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
         Left            =   4100
         TabIndex        =   64
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkordN2 
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
         Left            =   2700
         TabIndex        =   62
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkordN1 
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
         Left            =   1300
         TabIndex        =   61
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov8 
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
         Left            =   11300
         TabIndex        =   49
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr8 
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
         Left            =   11300
         TabIndex        =   48
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
         Left            =   11300
         TabIndex        =   47
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov7 
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
         Left            =   9700
         TabIndex        =   46
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr7 
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
         Left            =   9700
         TabIndex        =   45
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
         Left            =   9700
         TabIndex        =   44
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov5 
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
         Left            =   6900
         TabIndex        =   43
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr5 
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
         Left            =   6900
         TabIndex        =   42
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
         Left            =   6900
         TabIndex        =   41
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov4 
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
         Left            =   5500
         TabIndex        =   40
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr4 
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
         Left            =   5500
         TabIndex        =   39
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
         Left            =   5500
         TabIndex        =   38
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov2 
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
         Left            =   2700
         TabIndex        =   37
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr2 
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
         Left            =   2700
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
         Left            =   2700
         TabIndex        =   35
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorDov1 
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
         Left            =   1300
         TabIndex        =   34
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr1 
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
         Left            =   1300
         TabIndex        =   33
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
         Left            =   1300
         TabIndex        =   32
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
         Left            =   12700
         TabIndex        =   77
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
         Height          =   400
         Left            =   8300
         TabIndex        =   70
         Top             =   400
         Width           =   1200
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
         Left            =   4100
         TabIndex        =   63
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dN="
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
         TabIndex        =   60
         Top             =   2200
         Width           =   800
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dДов="
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
         TabIndex        =   31
         Top             =   2800
         Width           =   1000
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dПр="
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
         TabIndex        =   29
         Top             =   1000
         Width           =   700
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
         TabIndex        =   28
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
         Left            =   9700
         TabIndex        =   27
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
         Left            =   6900
         TabIndex        =   26
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
         Left            =   5500
         TabIndex        =   25
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
         Left            =   2700
         TabIndex        =   24
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
         Left            =   1300
         TabIndex        =   23
         Top             =   400
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   2600
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   18300
      Begin VB.TextBox pYr9 
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
         Left            =   14600
         TabIndex        =   59
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pXr9 
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
         Left            =   14600
         TabIndex        =   58
         Text            =   "0"
         Top             =   1000
         Width           =   1300
      End
      Begin VB.TextBox pYr6 
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
         Left            =   9500
         TabIndex        =   56
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pXr6 
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
         Left            =   9500
         TabIndex        =   55
         Text            =   "0"
         Top             =   1000
         Width           =   1300
      End
      Begin VB.TextBox pYr3 
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
         Left            =   4400
         TabIndex        =   53
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pXr3 
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
         Left            =   4400
         TabIndex        =   52
         Text            =   "0"
         Top             =   1000
         Width           =   1300
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "Решить"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   16500
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pYr8 
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
         Left            =   12900
         TabIndex        =   20
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pYr7 
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
         Left            =   11200
         TabIndex        =   19
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pYr5 
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
         Left            =   7800
         TabIndex        =   18
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pYr4 
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
         Left            =   6100
         TabIndex        =   17
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pYr2 
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
         Left            =   2700
         TabIndex        =   16
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pYr1 
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
         TabIndex        =   15
         Text            =   "0"
         Top             =   1600
         Width           =   1300
      End
      Begin VB.TextBox pXr8 
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
         Left            =   12900
         TabIndex        =   14
         Text            =   "0"
         Top             =   1000
         Width           =   1300
      End
      Begin VB.TextBox pXr7 
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
         Left            =   11200
         TabIndex        =   13
         Text            =   "0"
         Top             =   1000
         Width           =   1300
      End
      Begin VB.TextBox pXr5 
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
         Left            =   7800
         TabIndex        =   12
         Text            =   "0"
         Top             =   1000
         Width           =   1300
      End
      Begin VB.TextBox pXr4 
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
         Left            =   6100
         TabIndex        =   11
         Text            =   "0"
         Top             =   1000
         Width           =   1300
      End
      Begin VB.TextBox pXr2 
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
         Left            =   2700
         TabIndex        =   10
         Text            =   "0"
         Top             =   1000
         Width           =   1300
      End
      Begin VB.TextBox pXr1 
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
         TabIndex        =   9
         Text            =   "0"
         Top             =   1000
         Width           =   1300
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
         Left            =   14700
         TabIndex        =   57
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
         Left            =   9600
         TabIndex        =   54
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
         Left            =   4500
         TabIndex        =   51
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ур="
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
         TabIndex        =   8
         Top             =   1600
         Width           =   700
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Хр="
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
         TabIndex        =   7
         Top             =   1000
         Width           =   700
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
         Left            =   13000
         TabIndex        =   6
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
         Left            =   11300
         TabIndex        =   5
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
         Left            =   7900
         TabIndex        =   4
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
         Left            =   6200
         TabIndex        =   3
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
         Left            =   2800
         TabIndex        =   2
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
         Left            =   1100
         TabIndex        =   1
         Top             =   400
         Width           =   1200
      End
   End
End
Attribute VB_Name = "PrisXYform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Xr = pXr1:  Yr = pYr1
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pAks1X: Yb = bp6Oryd.pAks1Y: hb = bp6Oryd.pAks1h: dXtus = Shest6Oryd.pvdXtus1
        Dt = Shest6Oryd.pvDt1: Ygolt = Shest6Oryd.pvYgt1: snar = Shest6Oryd.pAks1Snar: zar = Shest6Oryd.pKal1Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD1 = dD: pvkorDov1 = dDov: pvkorPr1 = dPr: pvkordN1 = dN
End If

 Xr = pXr2:  Yr = pYr2
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pAks2X: Yb = bp6Oryd.pAks2Y: hb = bp6Oryd.pAks2h: dXtus = Shest6Oryd.pvdXtus2
        Dt = Shest6Oryd.pvDt2: Ygolt = Shest6Oryd.pvYgt2: snar = Shest6Oryd.pAks2Snar: zar = Shest6Oryd.pKal2Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD2 = dD: pvkorDov2 = dDov: pvkorPr2 = dPr: pvkordN2 = dN
End If

 Xr = pXr3:  Yr = pYr3
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pAks3X: Yb = bp6Oryd.pAks3Y: hb = bp6Oryd.pAks3h: dXtus = Shest6Oryd.pvdXtus3
        Dt = Shest6Oryd.pvDt3: Ygolt = Shest6Oryd.pvYgt3: snar = Shest6Oryd.pAks3Snar: zar = Shest6Oryd.pKal3Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD3 = dD: pvkorDov3 = dDov: pvkorPr3 = dPr: pvkordN3 = dN
End If

 Xr = pXr4:  Yr = pYr4
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pKal1X: Yb = bp6Oryd.pKal1Y: hb = bp6Oryd.pKal1h: dXtus = Shest6Oryd.pvdXtus4
        Dt = Shest6Oryd.pvDt4: Ygolt = Shest6Oryd.pvYgt4: snar = Shest6Oryd.pKal1Snar: zar = Shest6Oryd.pKal1Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD4 = dD: pvkorDov4 = dDov: pvkorPr4 = dPr: pvkordN4 = dN
End If

 Xr = pXr5:  Yr = pYr5
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pKal2X: Yb = bp6Oryd.pKal2Y: hb = bp6Oryd.pKal2h: dXtus = Shest6Oryd.pvdXtus5
        Dt = Shest6Oryd.pvDt5: Ygolt = Shest6Oryd.pvYgt5: snar = Shest6Oryd.pKal2Snar: zar = Shest6Oryd.pKal2Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD5 = dD: pvkorDov5 = dDov: pvkorPr5 = dPr: pvkordN5 = dN
End If

 Xr = pXr6:  Yr = pYr6
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pKal3X: Yb = bp6Oryd.pKal3Y: hb = bp6Oryd.pKal3h: dXtus = Shest6Oryd.pvdXtus6
        Dt = Shest6Oryd.pvDt6: Ygolt = Shest6Oryd.pvYgt6: snar = Shest6Oryd.pKal3Snar: zar = Shest6Oryd.pKal3Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD6 = dD: pvkorDov6 = dDov: pvkorPr6 = dPr: pvkordN6 = dN
End If

 Xr = pXr7:  Yr = pYr7
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pOsk1X: Yb = bp6Oryd.pOsk1Y: hb = bp6Oryd.pOsk1h: dXtus = Shest6Oryd.pvdXtus7
        Dt = Shest6Oryd.pvDt7: Ygolt = Shest6Oryd.pvYgt7: snar = Shest6Oryd.pOsk1Snar: zar = Shest6Oryd.pOsk1Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD7 = dD: pvkorDov7 = dDov: pvkorPr7 = dPr: pvkordN7 = dN
End If

 Xr = pXr8:  Yr = pYr8
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pOsk2X: Yb = bp6Oryd.pOsk2Y: hb = bp6Oryd.pOsk2h: dXtus = Shest6Oryd.pvdXtus8
        Dt = Shest6Oryd.pvDt8: Ygolt = Shest6Oryd.pvYgt8: snar = Shest6Oryd.pOsk2Snar: zar = Shest6Oryd.pOsk2Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD8 = dD: pvkorDov8 = dDov: pvkorPr8 = dPr: pvkordN8 = dN
End If

 Xr = pXr9:  Yr = pYr9
 If Xr = 0 And Yr = 0 Then
    Else
        Xb = bp6Oryd.pOsk3X: Yb = bp6Oryd.pOsk3Y: hb = bp6Oryd.pOsk3h: dXtus = Shest6Oryd.pvdXtus9
        Dt = Shest6Oryd.pvDt9: Ygolt = Shest6Oryd.pvYgt9: snar = Shest6Oryd.pOsk3Snar: zar = Shest6Oryd.pOsk3Zar
        podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvkorD9 = dD: pvkorDov9 = dDov: pvkorPr9 = dPr: pvkordN9 = dN
End If
End Sub

Private Sub Command2_Click()
PrisXYform.Hide
End Sub

Function podRASCHETPRIST6orXY(ByVal Xr As Single, ByVal Yr As Single, ByVal Xb As Single, ByVal Yb As Single, ByVal hb As Single, ByVal dXtus As Single, ByVal Dt As Single, ByVal Ygolt As Single, ByVal snar As String, ByVal zar As String, dD, dDov, dPr, dN) As Single
        dxr = Xr - Xb
        dyr = Yr - Yb
 Dtr = Sqr(dxr ^ 2 + dyr ^ 2)
 Ar = Abs(Atn(dyr / (dxr + 0.1)) / 3.141592 * 30) * 100
 If dxr > 0 And dyr > 0 Then Ygoltr = Int(Ar)
 If dxr < 0 And dyr > 0 Then Ygoltr = Int(3000 - Ar)
 If dxr < 0 And dyr < 0 Then Ygoltr = Int(3000 + Ar)
 If dxr > 0 And dyr < 0 Then Ygoltr = Int(6000 - Ar)
  dD = Round(Dt - Dtr)
 dDov = Round(Ygolt - Ygoltr)
 dPr = Round(dD / (dXtus + 0.01))
 If snar = "О-13" Or snar = "Ш2" Then
    dN = poldN(snar, zar, Dtr)
    dN = Round(dPr * dN)
 Else
    dN = 0
 End If
End Function
Function poldN(ByVal snar As String, ByVal zar As String, ByVal dalT As Single) As Single
Dim ta(20) As Single
Dim Yn As Single
If snar = "О -13" Then
            If zar = "Полн" Then
                Open App.Path & "\O13\polnuy" For Input As #1
            ElseIf zar = "Перв" Then
                Open App.Path & "\O13\pervuy" For Input As #1
            ElseIf zar = "Втор" Then
                Open App.Path & "\O13\vtoroi" For Input As #1
            ElseIf zar = "Трет" Then
                Open App.Path & "\O13\tretiy" For Input As #1
            ElseIf zar = "Четверт" Then
                Open App.Path & "\O13\chetvertuy" For Input As #1
            ElseIf zar = "Пятый" Then
                Open App.Path & "\O13\piatuy" For Input As #1
            ElseIf zar = "Шестой" Then
                Open App.Path & "\O13\shestoi" For Input As #1
            Else
                Open App.Path & "\O13\polnuy" For Input As #1
        End If
10:   If EOF(1) Then
            Yn = ta(5)
             GoTo 20
         Else
           Input #1, ta(1), ta(2), ta(3), ta(4), ta(5), ta(6), ta(7), ta(8), ta(9), ta(10), ta(11), ta(12), ta(13), ta(14), ta(15)
           If ta(1) > dalT Then
                Yn = ta(5)
                GoTo 20
           Else
                GoTo 10
           End If
         End If
20:   Close #1
poldN = (dalT + 0.001) / 1000 / (Yn + 0.001)
Else
    If zar = "Полн" Then
        Open App.Path & "\3SH-P" For Input As #1
     Else
        Open App.Path & "\3SH-Y" For Input As #1
    End If
30:   If EOF(1) Then
            Yn = ta(4)
            GoTo 40
        Else
              Input #1, ta(1), ta(2), ta(3), ta(4), ta(5), ta(6), ta(7), ta(8)
              If ta(1) > dalT Then
                Yn = ta(4)
                GoTo 40
              Else
                GoTo 30
              End If
        End If
40: Close #1
    poldN = Yn
End If
End Function

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

Private Sub pXr1_Click()
pXr1.Text = ""
End Sub
Private Sub pXr2_Click()
pXr2.Text = ""
End Sub
Private Sub pXr3_Click()
pXr3.Text = ""
End Sub
Private Sub pXr4_Click()
pXr4.Text = ""
End Sub
Private Sub pXr5_Click()
pXr5.Text = ""
End Sub
Private Sub pXr6_Click()
pXr6.Text = ""
End Sub
Private Sub pXr7_Click()
pXr7.Text = ""
End Sub
Private Sub pXr8_Click()
pXr8.Text = ""
End Sub
Private Sub pXr9_Click()
pXr9.Text = ""
End Sub

Private Sub pXr1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr1.Text = ""
pYr1.SetFocus
End If
End Sub
Private Sub pXr2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr2.Text = ""
pYr2.SetFocus
End If
End Sub
Private Sub pXr3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr3.Text = ""
pYr3.SetFocus
End If
End Sub
Private Sub pXr4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr4.Text = ""
pYr4.SetFocus
End If
End Sub
Private Sub pXr5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr5.Text = ""
pYr5.SetFocus
End If
End Sub
Private Sub pXr6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr6.Text = ""
pYr6.SetFocus
End If
End Sub
Private Sub pXr7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr7.Text = ""
pYr7.SetFocus
End If
End Sub
Private Sub pXr8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr8.Text = ""
pYr8.SetFocus
End If
End Sub
Private Sub pXr9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYr9.Text = ""
pYr9.SetFocus
End If
End Sub

