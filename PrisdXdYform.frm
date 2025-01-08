VERSION 5.00
Begin VB.Form PrisdXdYform 
   Caption         =   "Пристрелка dX, dY"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   17190
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
      Height          =   1200
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4800
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Корректуры"
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
      Left            =   2100
      TabIndex        =   27
      Top             =   3000
      Width           =   14900
      Begin VB.TextBox pvdDov9 
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
         Left            =   11600
         TabIndex        =   85
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
         Left            =   11600
         TabIndex        =   84
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvdPr9 
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
         Left            =   11600
         TabIndex        =   83
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD9 
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
         Height          =   480
         Left            =   11600
         TabIndex        =   82
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
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
         Left            =   10300
         TabIndex        =   81
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
         Left            =   9000
         TabIndex        =   80
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvdDov6 
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
         Left            =   7700
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
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   7700
         TabIndex        =   78
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvdPr6 
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
         Left            =   7700
         TabIndex        =   77
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD6 
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
         Height          =   480
         Left            =   7700
         TabIndex        =   76
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
         Left            =   6400
         TabIndex        =   74
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
         Left            =   5100
         TabIndex        =   73
         Text            =   "0"
         Top             =   2200
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
         Left            =   2600
         TabIndex        =   72
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
         Left            =   1300
         TabIndex        =   71
         Text            =   "0"
         Top             =   2200
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
         Left            =   3900
         TabIndex        =   69
         Text            =   "0"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox pvdDov3 
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
         Left            =   3900
         TabIndex        =   68
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdPr3 
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
         Left            =   3900
         TabIndex        =   67
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD3 
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
         Height          =   480
         Left            =   3900
         TabIndex        =   66
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvdDov8 
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
         Left            =   10300
         TabIndex        =   54
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdPr8 
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
         Left            =   10300
         TabIndex        =   53
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD8 
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
         Left            =   10300
         TabIndex        =   52
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvdDov7 
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
         Left            =   9000
         TabIndex        =   51
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdPr7 
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
         Left            =   9000
         TabIndex        =   50
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD7 
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
         Left            =   9000
         TabIndex        =   49
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvdDov5 
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
         Left            =   6400
         TabIndex        =   48
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdPr5 
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
         Left            =   6400
         TabIndex        =   47
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD5 
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
         Left            =   6400
         TabIndex        =   46
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvdDov4 
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
         Left            =   5100
         TabIndex        =   45
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdPr4 
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
         Left            =   5100
         TabIndex        =   44
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD4 
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
         Left            =   5100
         TabIndex        =   43
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvdDov2 
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
         Left            =   2600
         TabIndex        =   42
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdPr2 
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
         Height          =   450
         Left            =   2600
         TabIndex        =   41
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD2 
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
         Left            =   2600
         TabIndex        =   40
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvdDov1 
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
         Left            =   1300
         TabIndex        =   39
         Text            =   "0"
         Top             =   2800
         Width           =   1000
      End
      Begin VB.TextBox pvdPr1 
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
         Height          =   450
         Left            =   1300
         TabIndex        =   38
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pvdD1 
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
         TabIndex        =   37
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
         Left            =   11600
         TabIndex        =   86
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
         ForeColor       =   &H00404080&
         Height          =   400
         Left            =   7700
         TabIndex        =   75
         Top             =   400
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
         Height          =   405
         Left            =   120
         TabIndex        =   70
         Top             =   2205
         Width           =   795
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
         Left            =   3900
         TabIndex        =   65
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label19 
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
         TabIndex        =   36
         Top             =   2800
         Width           =   1000
      End
      Begin VB.Label Label18 
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
         Height          =   405
         Left            =   120
         TabIndex        =   35
         Top             =   1600
         Width           =   795
      End
      Begin VB.Label Label17 
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
         TabIndex        =   34
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
         Left            =   10300
         TabIndex        =   33
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
         Left            =   9000
         TabIndex        =   32
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
         Left            =   6450
         TabIndex        =   31
         Top             =   400
         Width           =   1000
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
         Left            =   5100
         TabIndex        =   30
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
         Left            =   2600
         TabIndex        =   29
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
         TabIndex        =   28
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
      Left            =   2200
      TabIndex        =   5
      Top             =   100
      Width           =   14850
      Begin VB.TextBox pdY9 
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
         Left            =   11600
         TabIndex        =   64
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX9 
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
         Left            =   11600
         TabIndex        =   63
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdY6 
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
         Left            =   7700
         TabIndex        =   61
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX6 
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
         Left            =   7700
         TabIndex        =   60
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdY3 
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
         Left            =   3800
         TabIndex        =   58
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX3 
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
         Left            =   3800
         TabIndex        =   57
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
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
         Height          =   1000
         Left            =   13300
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1000
         Width           =   1300
      End
      Begin VB.TextBox pdY8 
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
         Left            =   10300
         TabIndex        =   25
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX8 
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
         Left            =   10300
         TabIndex        =   24
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdY7 
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
         Left            =   9000
         TabIndex        =   23
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX7 
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
         Left            =   9000
         TabIndex        =   22
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdY5 
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
         Left            =   6400
         TabIndex        =   21
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX5 
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
         Left            =   6400
         TabIndex        =   20
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdY4 
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
         Left            =   5100
         TabIndex        =   19
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX4 
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
         Left            =   5100
         TabIndex        =   18
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdY2 
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
         TabIndex        =   17
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX2 
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
         TabIndex        =   16
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pdY1 
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
         Left            =   1200
         TabIndex        =   15
         Text            =   "0"
         Top             =   1600
         Width           =   1000
      End
      Begin VB.TextBox pdX1 
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
         Left            =   1200
         TabIndex        =   14
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
         Left            =   11600
         TabIndex        =   62
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
         ForeColor       =   &H00404080&
         Height          =   400
         Left            =   7750
         TabIndex        =   59
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
         ForeColor       =   &H00004080&
         Height          =   400
         Left            =   3800
         TabIndex        =   56
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dY="
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
         TabIndex        =   13
         Top             =   1600
         Width           =   700
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dX="
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
         TabIndex        =   12
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
         Left            =   10300
         TabIndex        =   11
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
         Left            =   9000
         TabIndex        =   10
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
         Left            =   6450
         TabIndex        =   9
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
         Left            =   5100
         TabIndex        =   8
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
         TabIndex        =   7
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
         Left            =   1200
         TabIndex        =   6
         Top             =   400
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   2700
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   1935
      Begin VB.OptionButton pkagdomy 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   2200
         Width           =   375
      End
      Begin VB.OptionButton pvsem 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   1000
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Каждому"
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
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Всем"
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
         Left            =   360
         TabIndex        =   1
         Top             =   400
         Width           =   800
      End
   End
End
Attribute VB_Name = "PrisdXdYform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xr As Single, Yr As Single, xc As Single, yc As Single
   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc1: yc = Zelkagdomy6.pYc1
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX1 + xc: Yr = pdY1 + yc
 If pdX1 = 0 And pdY1 = 0 Then
    Else
        Xb = bp6Oryd.pAks1X: Yb = bp6Oryd.pAks1Y: hb = bp6Oryd.pAks1h: dXtus = Shest6Oryd.pvdXtus1
        Dt = Shest6Oryd.pvDt1: Ygolt = Shest6Oryd.pvYgt1: snar = Shest6Oryd.pAks1Snar: zar = Shest6Oryd.pAks1Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD1 = dD: pvdDov1 = dDov: pvdPr1 = dPr: pvdN1 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc2: yc = Zelkagdomy6.pYc2
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX2 + xc: Yr = pdY2 + yc
 If pdX2 = 0 And pdY2 = 0 Then
    Else
        Xb = bp6Oryd.pAks2X: Yb = bp6Oryd.pAks2Y: hb = bp6Oryd.pAks2h: dXtus = Shest6Oryd.pvdXtus2
        Dt = Shest6Oryd.pvDt2: Ygolt = Shest6Oryd.pvYgt2: snar = Shest6Oryd.pAks2Snar: zar = Shest6Oryd.pAks2Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD2 = dD: pvdDov2 = dDov: pvdPr2 = dPr: pvdN2 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc3: yc = Zelkagdomy6.pYc3
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX3 + xc: Yr = pdY3 + yc
 If pdX3 = 0 And pdY3 = 0 Then
    Else
        Xb = bp6Oryd.pAks3X: Yb = bp6Oryd.pAks3Y: hb = bp6Oryd.pAks3h: dXtus = Shest6Oryd.pvdXtus3
        Dt = Shest6Oryd.pvDt3: Ygolt = Shest6Oryd.pvYgt3: snar = Shest6Oryd.pAks3Snar: zar = Shest6Oryd.pAks3Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD3 = dD: pvdDov3 = dDov: pvdPr3 = dPr: pvdN3 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc4: yc = Zelkagdomy6.pYc4
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX4 + xc: Yr = pdY4 + yc
 If pdX4 = 0 And pdY4 = 0 Then
    Else
        Xb = bp6Oryd.pKal1X: Yb = bp6Oryd.pKal1Y: hb = bp6Oryd.pKal1h: dXtus = Shest6Oryd.pvdXtus4
        Dt = Shest6Oryd.pvDt4: Ygolt = Shest6Oryd.pvYgt4: snar = Shest6Oryd.pKal1Snar: zar = Shest6Oryd.pKal1Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD4 = dD: pvdDov4 = dDov: pvdPr4 = dPr: pvdN4 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc5: yc = Zelkagdomy6.pYc5
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX5 + xc: Yr = pdY5 + yc
 If pdX5 = 0 And pdY5 = 0 Then
    Else
        Xb = bp6Oryd.pKal2X: Yb = bp6Oryd.pKal2Y: hb = bp6Oryd.pKal2h: dXtus = Shest6Oryd.pvdXtus5
        Dt = Shest6Oryd.pvDt5: Ygolt = Shest6Oryd.pvYgt5: snar = Shest6Oryd.pKal2Snar: zar = Shest6Oryd.pKal2Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD5 = dD: pvdDov5 = dDov: pvdPr5 = dPr: pvdN5 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc6: yc = Zelkagdomy6.pYc6
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX6 + xc: Yr = pdY6 + yc
 If pdX6 = 0 And pdY6 = 0 Then
    Else
        Xb = bp6Oryd.pKal3X: Yb = bp6Oryd.pKal3Y: hb = bp6Oryd.pKal3h: dXtus = Shest6Oryd.pvdXtus6
        Dt = Shest6Oryd.pvDt6: Ygolt = Shest6Oryd.pvYgt6: snar = Shest6Oryd.pKal3Snar: zar = Shest6Oryd.pKal3Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD6 = dD: pvdDov6 = dDov: pvdPr6 = dPr: pvdN6 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc7: yc = Zelkagdomy6.pYc7
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX7 + xc: Yr = pdY7 + yc
 If pdX7 = 0 And pdY7 = 0 Then
    Else
        Xb = bp6Oryd.pOsk1X: Yb = bp6Oryd.pOsk1Y: hb = bp6Oryd.pOsk1h: dXtus = Shest6Oryd.pvdXtus7
        Dt = Shest6Oryd.pvDt7: Ygolt = Shest6Oryd.pvYgt7: snar = Shest6Oryd.pOsk1Snar: zar = Shest6Oryd.pOsk1Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD7 = dD: pvdDov7 = dDov: pvdPr7 = dPr: pvdN7 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc8: yc = Zelkagdomy6.pYc8
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX8 + xc: Yr = pdY8 + yc
 If pdX8 = 0 And pdY8 = 0 Then
    Else
        Xb = bp6Oryd.pOsk2X: Yb = bp6Oryd.pOsk2Y: hb = bp6Oryd.pOsk2h: dXtus = Shest6Oryd.pvdXtus8
        Dt = Shest6Oryd.pvDt8: Ygolt = Shest6Oryd.pvYgt8: snar = Shest6Oryd.pOsk2Snar: zar = Shest6Oryd.pOsk2Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD8 = dD: pvdDov8 = dDov: pvdPr8 = dPr: pvdN8 = dN
End If

   If pvsem = True Then
        xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
        ElseIf pkagdomy = True Then
            xc = Zelkagdomy6.pXc9: yc = Zelkagdomy6.pYc9
            Else
                       xc = Shest6Oryd.pXc: yc = Shest6Oryd.pYc
End If
 Xr = pdX9 + xc: Yr = pdY9 + yc
 If pdX9 = 0 And pdY9 = 0 Then
    Else
        Xb = bp6Oryd.pOsk3X: Yb = bp6Oryd.pOsk3Y: hb = bp6Oryd.pOsk3h: dXtus = Shest6Oryd.pvdXtus9
        Dt = Shest6Oryd.pvDt9: Ygolt = Shest6Oryd.pvYgt9:: snar = Shest6Oryd.pOsk3Snar: zar = Shest6Oryd.pOsk3Zar
        PrisXYform.podRASCHETPRIST6orXY Xr, Yr, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dN
        pvdD9 = dD: pvdDov9 = dDov: pvdPr9 = dPr: pvdN9 = dN
End If

End Sub

Private Sub Command2_Click()
PrisdXdYform.Hide
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

Private Sub pdX1_Click()
pdX1.Text = ""
End Sub
Private Sub pdX2_Click()
pdX2.Text = ""
End Sub
Private Sub pdX3_Click()
pdX3.Text = ""
End Sub
Private Sub pdX4_Click()
pdX4.Text = ""
End Sub
Private Sub pdX5_Click()
pdX5.Text = ""
End Sub
Private Sub pdX6_Click()
pdX6.Text = ""
End Sub
Private Sub pdX7_Click()
pdX7.Text = ""
End Sub
Private Sub pdX8_Click()
pdX8.Text = ""
End Sub
Private Sub pdX9_Click()
pdX9.Text = ""
End Sub

Private Sub pdX1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY1.Text = ""
pdY1.SetFocus
End If
End Sub
Private Sub pdX2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY2.Text = ""
pdY2.SetFocus
End If
End Sub
Private Sub pdX3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY3.Text = ""
pdY3.SetFocus
End If
End Sub
Private Sub pdX4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY4.Text = ""
pdY4.SetFocus
End If
End Sub
Private Sub pdX5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY5.Text = ""
pdY5.SetFocus
End If
End Sub
Private Sub pdX6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY6.Text = ""
pdY6.SetFocus
End If
End Sub
Private Sub pdX7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY7.Text = ""
pdY7.SetFocus
End If
End Sub
Private Sub pdX8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY8.Text = ""
pdY8.SetFocus
End If
End Sub
Private Sub pdX9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pdY9.Text = ""
pdY9.SetFocus
End If
End Sub

