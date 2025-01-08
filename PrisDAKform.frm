VERSION 5.00
Begin VB.Form PrisDAKform 
   Caption         =   "Пристрелка ДАК"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   18405
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6000
      Width           =   1455
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
      Height          =   4095
      Left            =   2500
      TabIndex        =   25
      Top             =   3100
      Width           =   15700
      Begin VB.TextBox pvkorNap9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   540
         Left            =   12300
         TabIndex        =   84
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkordN9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   540
         Left            =   12300
         TabIndex        =   83
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr9 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   540
         Left            =   12300
         TabIndex        =   82
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorD9 
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
         Height          =   500
         Left            =   12300
         TabIndex        =   81
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorNap6 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   500
         Left            =   8100
         TabIndex        =   79
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkordN6 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   500
         Left            =   8100
         TabIndex        =   78
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr6 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   500
         Left            =   8100
         TabIndex        =   77
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorD6 
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
         Height          =   500
         Left            =   8100
         TabIndex        =   76
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkordN8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   500
         Left            =   10900
         TabIndex        =   74
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkordN7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   500
         Left            =   9500
         TabIndex        =   73
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkordN5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   480
         Left            =   6700
         TabIndex        =   72
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkordN4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   5300
         TabIndex        =   71
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkordN3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
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
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkordN2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   2500
         TabIndex        =   68
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkordN1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1100
         TabIndex        =   67
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pvkorNap3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   540
         Left            =   3900
         TabIndex        =   66
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr3 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   3900
         TabIndex        =   65
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorD3 
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
         TabIndex        =   64
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   540
         Left            =   10900
         TabIndex        =   53
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   540
         Left            =   9500
         TabIndex        =   52
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   540
         Left            =   6700
         TabIndex        =   51
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   540
         Left            =   5300
         TabIndex        =   50
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   2500
         TabIndex        =   49
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorPr1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1100
         TabIndex        =   48
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pvkorNap8 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   540
         Left            =   10900
         TabIndex        =   45
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkorD8 
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
         Left            =   10900
         TabIndex        =   44
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorNap7 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   540
         Left            =   9500
         TabIndex        =   43
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkorD7 
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
         Left            =   9500
         TabIndex        =   42
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorNap5 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   540
         Left            =   6700
         TabIndex        =   41
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkorD5 
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
         Left            =   6700
         TabIndex        =   40
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorNap4 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   540
         Left            =   5300
         TabIndex        =   39
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkorD4 
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
         Left            =   5300
         TabIndex        =   38
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorNap2 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   2500
         TabIndex        =   37
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkorD2 
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
         Left            =   2500
         TabIndex        =   36
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pvkorNap1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1100
         TabIndex        =   35
         Text            =   "0"
         Top             =   3100
         Width           =   1000
      End
      Begin VB.TextBox pvkorD1 
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
         Left            =   1100
         TabIndex        =   34
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
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   12300
         TabIndex        =   80
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
         Left            =   8100
         TabIndex        =   75
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label23 
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
         TabIndex        =   70
         Top             =   2400
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
         Left            =   3900
         TabIndex        =   63
         Top             =   400
         Width           =   1200
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
         Height          =   400
         Left            =   100
         TabIndex        =   47
         Top             =   1700
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
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   10900
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
         ForeColor       =   &H00800000&
         Height          =   400
         Left            =   9500
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
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   6700
         TabIndex        =   31
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
         Left            =   5300
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
         Left            =   2500
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
         Left            =   1100
         TabIndex        =   28
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label11 
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
         TabIndex        =   27
         Top             =   3100
         Width           =   1000
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "dД="
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
         TabIndex        =   26
         Top             =   1005
         Width           =   600
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
      Height          =   2800
      Left            =   2800
      TabIndex        =   3
      Top             =   100
      Width           =   15400
      Begin VB.TextBox pDr9 
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
         Left            =   12000
         TabIndex        =   62
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr9 
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
         Left            =   12000
         TabIndex        =   61
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pDr6 
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
         Left            =   7800
         TabIndex        =   59
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr6 
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
         Left            =   7800
         TabIndex        =   58
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pDr3 
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
         Left            =   3600
         TabIndex        =   56
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr3 
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
         Left            =   3600
         TabIndex        =   55
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
         Left            =   13500
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1000
         Width           =   1400
      End
      Begin VB.TextBox pDr8 
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
         Left            =   10600
         TabIndex        =   23
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr8 
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
         Left            =   10600
         TabIndex        =   22
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pDr7 
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
         Left            =   9200
         TabIndex        =   21
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr7 
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
         Left            =   9200
         TabIndex        =   20
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pDr5 
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
         Left            =   6400
         TabIndex        =   19
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr5 
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
         Left            =   6400
         TabIndex        =   18
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pDr4 
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
         Left            =   5000
         TabIndex        =   17
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr4 
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
         Left            =   5000
         TabIndex        =   16
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pDr2 
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
         Left            =   2200
         TabIndex        =   15
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr2 
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
         Left            =   2200
         TabIndex        =   14
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pDr1 
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
         TabIndex        =   8
         Text            =   "0"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox pAr1 
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
         ForeColor       =   &H00008080&
         Height          =   400
         Left            =   12000
         TabIndex        =   60
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
         Height          =   375
         Left            =   7800
         TabIndex        =   57
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
         Left            =   3600
         TabIndex        =   54
         Top             =   400
         Width           =   1200
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
         ForeColor       =   &H00800080&
         Height          =   400
         Left            =   10600
         TabIndex        =   13
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
         ForeColor       =   &H00C00000&
         Height          =   400
         Left            =   9200
         TabIndex        =   12
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
         ForeColor       =   &H00008000&
         Height          =   400
         Left            =   6400
         TabIndex        =   11
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
         Left            =   5000
         TabIndex        =   10
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
         Left            =   2100
         TabIndex        =   9
         Top             =   400
         Width           =   1200
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Д="
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
         TabIndex        =   6
         Top             =   1700
         Width           =   600
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
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
         TabIndex        =   5
         Top             =   1000
         Width           =   600
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
         Left            =   800
         TabIndex        =   4
         Top             =   400
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "КНП №"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   120
      TabIndex        =   0
      Top             =   100
      Width           =   2500
      Begin VB.ComboBox pNKp 
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
         ItemData        =   "PrisDAKform.frx":0000
         Left            =   1000
         List            =   "PrisDAKform.frx":0013
         TabIndex        =   2
         Text            =   "1"
         Top             =   600
         Width           =   700
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "НП="
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
         Top             =   600
         Width           =   700
      End
   End
End
Attribute VB_Name = "PrisDAKform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Arr As Single, Dr As Single, Mr As Single, nkp As Single, nb As Single, Xkp As Single, Ykp As Single, hkp As Single, xc As Single, yc As Single, hc As Single
  nkp = pNKp
   If nkp = 1 Then Xkp = BP.pXkp1: Ykp = BP.pYkp1: hkp = BP.phkp1
   If nkp = 2 Then Xkp = BP.pXkp2: Ykp = BP.pYkp2: hkp = BP.phkp2
   If nkp = 3 Then Xkp = BP.pXkp3: Ykp = BP.pYkp3: hkp = BP.phkp3
   If nkp = 4 Then Xkp = BP.pXkp4: Ykp = BP.pYkp4: hkp = BP.phkp4
   If nkp = 5 Then Xkp = BP.pXkp5: Ykp = BP.pYkp5: hkp = BP.phkp5
   
 Dr = pDr1:  Arr = pAr1
 If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pAks1X: Yb = bp6Oryd.pAks1Y: hb = bp6Oryd.pAks1h: dXtus = Shest6Oryd.pvdXtus1: Dt = Shest6Oryd.pvDt1
        Ygolt = Shest6Oryd.pvYgt1: snar = Shest6Oryd.pAks1Snar: zar = Shest6Oryd.pAks1Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD1 = dD: pvkorNap1 = dDov: pvkorPr1 = dPr: pvkordN1 = dN
End If

 Dr = pDr2:  Arr = pAr2
  If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pAks2X: Yb = bp6Oryd.pAks2Y: hb = bp6Oryd.pAks2h
        dXtus = Shest6Oryd.pvdXtus2: Dt = Shest6Oryd.pvDt2: Ygolt = Shest6Oryd.pvYgt2: snar = Shest6Oryd.pAks2Snar: zar = Shest6Oryd.pAks2Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD2 = dD: pvkorNap2 = dDov: pvkorPr2 = dPr: pvkordN2 = dN
End If

 Dr = pDr3:  Arr = pAr3
  If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pAks3X: Yb = bp6Oryd.pAks3Y: hb = bp6Oryd.pAks3h
        dXtus = Shest6Oryd.pvdXtus3: Dt = Shest6Oryd.pvDt3: Ygolt = Shest6Oryd.pvYgt3: snar = Shest6Oryd.pAks3Snar: zar = Shest6Oryd.pAks3Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD3 = dD: pvkorNap3 = dDov: pvkorPr3 = dPr: pvkordN3 = dN
End If

 Dr = pDr4:  Arr = pAr4
  If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pKal1X: Yb = bp6Oryd.pKal1Y: hb = bp6Oryd.pKal1h
        dXtus = Shest6Oryd.pvdXtus4: Dt = Shest6Oryd.pvDt4: Ygolt = Shest6Oryd.pvYgt4: snar = Shest6Oryd.pKal1Snar: zar = Shest6Oryd.pKal1Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD4 = dD: pvkorNap4 = dDov: pvkorPr4 = dPr: pvkordN4 = dN
End If

 Dr = pDr5:  Arr = pAr5
  If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pKal2X: Yb = bp6Oryd.pKal2Y: hb = bp6Oryd.pKal2h
        dXtus = Shest6Oryd.pvdXtus5: Dt = Shest6Oryd.pvDt5: Ygolt = Shest6Oryd.pvYgt5: snar = Shest6Oryd.pKal2Snar: zar = Shest6Oryd.pKal2Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD5 = dD: pvkorNap5 = dDov: pvkorPr5 = dPr: pvkordN5 = dN
End If

 Dr = pDr6:  Arr = pAr6
  If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pKal3X: Yb = bp6Oryd.pKal3Y: hb = bp6Oryd.pKal3h
        dXtus = Shest6Oryd.pvdXtus6: Dt = Shest6Oryd.pvDt6: Ygolt = Shest6Oryd.pvYgt6: snar = Shest6Oryd.pKal3Snar: zar = Shest6Oryd.pKal3Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD6 = dD: pvkorNap6 = dDov: pvkorPr6 = dPr: pvkordN6 = dN
End If

 Dr = pDr7:  Arr = pAr7
  If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pOsk1X: Yb = bp6Oryd.pOsk1Y: hb = bp6Oryd.pOsk1h
        dXtus = Shest6Oryd.pvdXtus7: Dt = Shest6Oryd.pvDt7: Ygolt = Shest6Oryd.pvYgt7: snar = Shest6Oryd.pOsk1Snar: zar = Shest6Oryd.pOsk1Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD7 = dD: pvkorNap7 = dDov: pvkorPr7 = dPr: pvkordN7 = dN
End If

 Dr = pDr8:  Arr = pAr8
  If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pOsk2X: Yb = bp6Oryd.pOsk2Y: hb = bp6Oryd.pOsk2h
        dXtus = Shest6Oryd.pvdXtus8: Dt = Shest6Oryd.pvDt8: Ygolt = Shest6Oryd.pvYgt8: snar = Shest6Oryd.pOsk2Snar: zar = Shest6Oryd.pOsk2Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD8 = dD: pvkorNap8 = dDov: pvkorPr8 = dPr: pvkordN8 = dN
End If

 Dr = pDr9:  Arr = pAr9
  If Dr = 0 And Arr = 0 Then
    Else
        Xb = bp6Oryd.pOsk3X: Yb = bp6Oryd.pOsk3Y: hb = bp6Oryd.pOsk3h
        dXtus = Shest6Oryd.pvdXtus9: Dt = Shest6Oryd.pvDt9: Ygolt = Shest6Oryd.pvYgt9: snar = Shest6Oryd.pOsk3Snar: zar = Shest6Oryd.pOsk3Zar
        podRASCHETPRIST6or Arr, Dr, Xkp, Ykp, Mr, hkp, Xb, Yb, hb, dXtus, Dt, Ygolt, snar, zar, dD, dDov, dPr, dYrr, dN
        pvkorD9 = dD: pvkorNap9 = dDov: pvkorPr9 = dPr: pvkordN9 = dN
End If

End Sub

Private Sub Command2_Click()
PrisDAKform.Hide
End Sub

Sub podRASCHETPRIST6or(ByVal Arr As Single, ByVal Dr As Single, ByVal Xkp As Single, ByVal Ykp As Single, ByVal Mr As Single, ByVal hkp As Single, ByVal Xb As Single, ByVal Yb As Single, ByVal hb As Single, ByVal dXtus As Single, ByVal Dt As Single, ByVal Ygolt As Single, ByVal snar As String, ByVal zar As String, dD, dDov, dPr, dYrr, dN)
152 Xr = Cos(Arr / 100 * 6 * 3.141592 / 180) * Dr + Xkp
 Yr = Sin(Arr / 100 * 6 * 3.141592 / 180) * Dr + Ykp
 hr = (Mr * (Dr * 0.001)) * 1.05 + hkp
 hc = Shest6Oryd.phc
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
dYrr = Round(((hc - hr + 0.001) / ((Dt + 0.001) * 0.001)) * 0.95)
dNtus = PrisXYform.poldN(snar, zar, Dt)
dN = Round(dPr * (dNtus + 0.001))
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

Private Sub pAr1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 pDr1.Text = ""
 pDr1.SetFocus
 End If
End Sub
Private Sub pAr2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 pDr2.Text = ""
 pDr2.SetFocus
 End If
End Sub
Private Sub pAr3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 pDr3.Text = ""
 pDr3.SetFocus
 End If
End Sub
Private Sub pAr4_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 pDr4.Text = ""
 pDr4.SetFocus
 End If
End Sub
Private Sub pAr5_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 pDr5.Text = ""
 pDr5.SetFocus
 End If
End Sub
Private Sub pAr6_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 pDr6.Text = ""
 pDr6.SetFocus
 End If
End Sub

