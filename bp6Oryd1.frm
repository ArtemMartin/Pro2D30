VERSION 5.00
Begin VB.Form bp6Oryd 
   Caption         =   "Боевой порядок  9 орудий"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   11265
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnZapisat 
      BackColor       =   &H00FF8080&
      Caption         =   "ЗАПИСАТЬ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   8900
      Width           =   5400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "ВЫХОД"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   100
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   8900
      Width           =   5400
   End
   Begin VB.Frame Frame1 
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
      Height          =   8655
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   11025
      Begin VB.TextBox pOr9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   480
         Left            =   100
         TabIndex        =   70
         Text            =   "_"
         Top             =   7400
         Width           =   1500
      End
      Begin VB.TextBox pOr8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   100
         TabIndex        =   69
         Text            =   "_"
         Top             =   6600
         Width           =   1500
      End
      Begin VB.TextBox pOr7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   100
         TabIndex        =   68
         Text            =   "_"
         Top             =   5800
         Width           =   1500
      End
      Begin VB.TextBox pOr6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   480
         Left            =   100
         TabIndex        =   67
         Text            =   "_"
         Top             =   5000
         Width           =   1500
      End
      Begin VB.TextBox pOr5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   480
         Left            =   100
         TabIndex        =   66
         Text            =   "_"
         Top             =   4200
         Width           =   1500
      End
      Begin VB.TextBox pOr4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   100
         TabIndex        =   65
         Text            =   "_"
         Top             =   3400
         Width           =   1500
      End
      Begin VB.TextBox pOr3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   480
         Left            =   100
         TabIndex        =   64
         Text            =   "_"
         Top             =   2600
         Width           =   1500
      End
      Begin VB.TextBox pOr2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   100
         TabIndex        =   63
         Text            =   "_"
         Top             =   1800
         Width           =   1500
      End
      Begin VB.TextBox pOr1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   100
         TabIndex        =   62
         Text            =   "_"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pOsk3V0 
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
         Left            =   9500
         TabIndex        =   60
         Text            =   "0"
         Top             =   7400
         Width           =   1000
      End
      Begin VB.TextBox pOsk3Tz 
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
         Left            =   8200
         TabIndex        =   59
         Text            =   "0"
         Top             =   7400
         Width           =   1000
      End
      Begin VB.TextBox pOsk3ON 
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
         Left            =   6900
         TabIndex        =   58
         Text            =   "0"
         Top             =   7400
         Width           =   1000
      End
      Begin VB.TextBox pOsk3h 
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
         Left            =   5600
         TabIndex        =   57
         Text            =   "0"
         Top             =   7400
         Width           =   1000
      End
      Begin VB.TextBox pOsk3Y 
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
         Left            =   3800
         TabIndex        =   56
         Text            =   "0"
         Top             =   7400
         Width           =   1500
      End
      Begin VB.TextBox pOsk3X 
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
         Top             =   7400
         Width           =   1500
      End
      Begin VB.TextBox pKal3V0 
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
         Left            =   9500
         TabIndex        =   54
         Text            =   "0"
         Top             =   5000
         Width           =   1000
      End
      Begin VB.TextBox pKal3Tz 
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
         Left            =   8200
         TabIndex        =   53
         Text            =   "0"
         Top             =   5000
         Width           =   1000
      End
      Begin VB.TextBox pKal3ON 
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
         Left            =   6900
         TabIndex        =   52
         Text            =   "0"
         Top             =   5000
         Width           =   1000
      End
      Begin VB.TextBox pKal3h 
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
         Left            =   5600
         TabIndex        =   51
         Text            =   "0"
         Top             =   5000
         Width           =   1000
      End
      Begin VB.TextBox pKal3Y 
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
         Left            =   3800
         TabIndex        =   50
         Text            =   "0"
         Top             =   5000
         Width           =   1500
      End
      Begin VB.TextBox pKal3X 
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
         TabIndex        =   49
         Text            =   "0"
         Top             =   5000
         Width           =   1500
      End
      Begin VB.TextBox pAks3V0 
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
         Left            =   9500
         TabIndex        =   48
         Text            =   "0"
         Top             =   2600
         Width           =   1000
      End
      Begin VB.TextBox pAks3Tz 
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
         Left            =   8200
         TabIndex        =   47
         Text            =   "0"
         Top             =   2600
         Width           =   1000
      End
      Begin VB.TextBox pAks3ON 
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
         Left            =   6900
         TabIndex        =   46
         Text            =   "0"
         Top             =   2600
         Width           =   1000
      End
      Begin VB.TextBox pAks3h 
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
         Left            =   5600
         TabIndex        =   45
         Text            =   "0"
         Top             =   2600
         Width           =   1000
      End
      Begin VB.TextBox pAks3Y 
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
         Left            =   3800
         TabIndex        =   44
         Text            =   "0"
         Top             =   2600
         Width           =   1500
      End
      Begin VB.TextBox pAks3X 
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
         Left            =   2000
         TabIndex        =   43
         Text            =   "0"
         Top             =   2600
         Width           =   1500
      End
      Begin VB.TextBox pAks1X 
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
         TabIndex        =   36
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pAks1Y 
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
         TabIndex        =   35
         Text            =   "0"
         Top             =   1000
         Width           =   1500
      End
      Begin VB.TextBox pAks1h 
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
         TabIndex        =   34
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pAks1ON 
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
         Left            =   6900
         TabIndex        =   33
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pAks2X 
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
         TabIndex        =   32
         Text            =   "0"
         Top             =   1800
         Width           =   1500
      End
      Begin VB.TextBox pAks2Y 
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
         Left            =   3800
         TabIndex        =   31
         Text            =   "0"
         Top             =   1800
         Width           =   1500
      End
      Begin VB.TextBox pAks2h 
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
         Left            =   5600
         TabIndex        =   30
         Text            =   "0"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox pAks2ON 
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
         Left            =   6900
         TabIndex        =   29
         Text            =   "0"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox pKal1X 
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
         TabIndex        =   28
         Text            =   "0"
         Top             =   3400
         Width           =   1500
      End
      Begin VB.TextBox pKal1Y 
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
         Left            =   3800
         TabIndex        =   27
         Text            =   "0"
         Top             =   3400
         Width           =   1500
      End
      Begin VB.TextBox pKal1h 
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
         TabIndex        =   26
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.TextBox pKal1ON 
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
         Left            =   6900
         TabIndex        =   25
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.TextBox pKal2X 
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
         TabIndex        =   24
         Text            =   "0"
         Top             =   4200
         Width           =   1500
      End
      Begin VB.TextBox pKal2Y 
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
         Left            =   3800
         TabIndex        =   23
         Text            =   "0"
         Top             =   4200
         Width           =   1500
      End
      Begin VB.TextBox pKal2h 
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
         Left            =   5600
         TabIndex        =   22
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pKal2ON 
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
         TabIndex        =   21
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pOsk1X 
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
         Left            =   2000
         TabIndex        =   20
         Text            =   "0"
         Top             =   5800
         Width           =   1500
      End
      Begin VB.TextBox pOsk2X 
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
         Left            =   2000
         TabIndex        =   19
         Text            =   "0"
         Top             =   6600
         Width           =   1500
      End
      Begin VB.TextBox pOsk1Y 
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
         Left            =   3800
         TabIndex        =   18
         Text            =   "0"
         Top             =   5800
         Width           =   1500
      End
      Begin VB.TextBox pOsk2Y 
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
         Left            =   3800
         TabIndex        =   17
         Text            =   "0"
         Top             =   6600
         Width           =   1500
      End
      Begin VB.TextBox pOsk1h 
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
         Left            =   5600
         TabIndex        =   16
         Text            =   "0"
         Top             =   5800
         Width           =   1000
      End
      Begin VB.TextBox pOsk2h 
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
         Left            =   5600
         TabIndex        =   15
         Text            =   "0"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.TextBox pOsk1ON 
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
         Left            =   6900
         TabIndex        =   14
         Text            =   "0"
         Top             =   5800
         Width           =   1000
      End
      Begin VB.TextBox pOsk2ON 
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
         Left            =   6900
         TabIndex        =   13
         Text            =   "0"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.TextBox pAks1Tz 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   8200
         TabIndex        =   12
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pAks2Tz 
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
         Height          =   500
         Left            =   8200
         TabIndex        =   11
         Text            =   "0"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox pKal1Tz 
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
         Height          =   500
         Left            =   8200
         TabIndex        =   10
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.TextBox pKal2Tz 
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
         Height          =   500
         Left            =   8200
         TabIndex        =   9
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pOsk1Tz 
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
         Height          =   500
         Left            =   8200
         TabIndex        =   8
         Text            =   "0"
         Top             =   5800
         Width           =   1000
      End
      Begin VB.TextBox pOsk2Tz 
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
         Height          =   500
         Left            =   8200
         TabIndex        =   7
         Text            =   "0"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.TextBox pAks1V0 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   9500
         TabIndex        =   6
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pAks2V0 
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
         Height          =   500
         Left            =   9500
         TabIndex        =   5
         Text            =   "0"
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox pKal1V0 
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
         Height          =   500
         Left            =   9500
         TabIndex        =   4
         Text            =   "0"
         Top             =   3400
         Width           =   1000
      End
      Begin VB.TextBox pKal2V0 
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
         Height          =   500
         Left            =   9500
         TabIndex        =   3
         Text            =   "0"
         Top             =   4200
         Width           =   1000
      End
      Begin VB.TextBox pOsk1V0 
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
         Height          =   500
         Left            =   9500
         TabIndex        =   2
         Text            =   "0"
         Top             =   5800
         Width           =   1000
      End
      Begin VB.TextBox pOsk2V0 
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
         Height          =   500
         Left            =   9500
         TabIndex        =   1
         Text            =   "0"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Позывной"
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
         Left            =   100
         TabIndex        =   71
         Top             =   400
         Width           =   1500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2500
         TabIndex        =   42
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   4400
         TabIndex        =   41
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5920
         TabIndex        =   40
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OH"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   7120
         TabIndex        =   39
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tz"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   8440
         TabIndex        =   38
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "V0"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   9880
         TabIndex        =   37
         Top             =   400
         Width           =   600
      End
   End
End
Attribute VB_Name = "bp6Oryd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnZapisat_Click()
Open App.Path & "\BP\bp9.txt" For Output As #1
Write #1, pOr1, pAks1X, pAks1Y, pAks1h, pAks1ON, pAks1Tz, pAks1V0
Write #1, pOr2, pAks2X, pAks2Y, pAks2h, pAks2ON, pAks2Tz, pAks2V0
Write #1, pOr3, pAks3X, pAks3Y, pAks3h, pAks3ON, pAks3Tz, pAks3V0
Write #1, pOr4, pKal1X, pKal1Y, pKal1h, pKal1ON, pKal1Tz, pKal1V0
Write #1, pOr5, pKal2X, pKal2Y, pKal2h, pKal2ON, pKal2Tz, pKal2V0
Write #1, pOr6, pKal3X, pKal3Y, pKal3h, pKal3ON, pKal3Tz, pKal3V0
Write #1, pOr7, pOsk1X, pOsk1Y, pOsk1h, pOsk1ON, pOsk1Tz, pOsk1V0
Write #1, pOr8, pOsk2X, pOsk2Y, pOsk2h, pOsk2ON, pOsk2Tz, pOsk2V0
Write #1, pOr9, pOsk3X, pOsk3Y, pOsk3h, pOsk3ON, pOsk3Tz, pOsk3V0
Close #1
End Sub

Private Sub Command1_Click()
bp6Oryd.Hide
End Sub

Private Sub Form_Load()
Dim t(10) As String
Open App.Path & "\BP\bp9.txt" For Input As #1
Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr1 = t(0): pAks1X = t(1): pAks1Y = t(2): pAks1h = t(3): pAks1ON = t(4): pAks1Tz = t(5): pAks1V0 = t(6)
Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr2 = t(0): pAks2X = t(1): pAks2Y = t(2): pAks2h = t(3): pAks2ON = t(4): pAks2Tz = t(5): pAks2V0 = t(6)
Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr3 = t(0): pAks3X = t(1): pAks3Y = t(2): pAks3h = t(3): pAks3ON = t(4): pAks3Tz = t(5): pAks3V0 = t(6)

Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr4 = t(0): pKal1X = t(1): pKal1Y = t(2): pKal1h = t(3): pKal1ON = t(4): pKal1Tz = t(5): pKal1V0 = t(6)
Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr5 = t(0): pKal2X = t(1): pKal2Y = t(2): pKal2h = t(3): pKal2ON = t(4): pKal2Tz = t(5): pKal2V0 = t(6)
Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr6 = t(0): pKal3X = t(1): pKal3Y = t(2): pKal3h = t(3): pKal3ON = t(4): pKal3Tz = t(5): pKal3V0 = t(6)

Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr7 = t(0): pOsk1X = t(1): pOsk1Y = t(2): pOsk1h = t(3): pOsk1ON = t(4): pOsk1Tz = t(5): pOsk1V0 = t(6)
Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr8 = t(0): pOsk2X = t(1): pOsk2Y = t(2): pOsk2h = t(3): pOsk2ON = t(4): pOsk2Tz = t(5): pOsk2V0 = t(6)
Input #1, t(0), t(1), t(2), t(3), t(4), t(5), t(6)
pOr9 = t(0): pOsk3X = t(1): pOsk3Y = t(2): pOsk3h = t(3): pOsk3ON = t(4): pOsk3Tz = t(5): pOsk3V0 = t(6)
Close #1
End Sub

Private Sub pAks1X_Click()
pAks1X.Text = ""
End Sub

Private Sub pAks1X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks1Y.Text = ""
    pAks1Y.SetFocus
    Else
End If
End Sub
Private Sub pAks1y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks1h.Text = ""
    pAks1h.SetFocus
    Else
End If
End Sub
Private Sub pAks1h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks1ON.Text = ""
    pAks1ON.SetFocus
    Else
End If
End Sub
Private Sub pAks1on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks1Tz.Text = ""
    pAks1Tz.SetFocus
    Else
End If
End Sub
Private Sub pAks1tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks1V0.Text = ""
    pAks1V0.SetFocus
    Else
End If
End Sub

Private Sub pAks2X_Click()
pAks2X.Text = ""
End Sub

Private Sub pAks2X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks2Y.Text = ""
    pAks2Y.SetFocus
    Else
End If
End Sub
Private Sub pAks2y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks2h.Text = ""
    pAks2h.SetFocus
    Else
End If
End Sub
Private Sub pAks2h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks2ON.Text = ""
    pAks2ON.SetFocus
    Else
End If
End Sub
Private Sub pAks2on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks2Tz.Text = ""
    pAks2Tz.SetFocus
    Else
End If
End Sub
Private Sub pAks2tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks2V0.Text = ""
    pAks2V0.SetFocus
    Else
End If
End Sub
Private Sub pAks3X_Click()
pAks3X.Text = ""
End Sub

Private Sub pAks3X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks3Y.Text = ""
    pAks3Y.SetFocus
    Else
End If
End Sub
Private Sub pAks3y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks3h.Text = ""
    pAks3h.SetFocus
    Else
End If
End Sub
Private Sub pAks3h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks3ON.Text = ""
    pAks3ON.SetFocus
    Else
End If
End Sub
Private Sub pAks3on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks3Tz.Text = ""
    pAks3Tz.SetFocus
    Else
End If
End Sub
Private Sub pAks3tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pAks3V0.Text = ""
    pAks3V0.SetFocus
    Else
End If
End Sub

Private Sub pKal1X_Click()
pKal1X.Text = ""
End Sub

Private Sub pKal1X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal1Y.Text = ""
    pKal1Y.SetFocus
    Else
End If
End Sub
Private Sub pKal1y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal1h.Text = ""
    pKal1h.SetFocus
    Else
End If
End Sub
Private Sub pKal1h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal1ON.Text = ""
    pKal1ON.SetFocus
    Else
End If
End Sub
Private Sub pKal1on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal1Tz.Text = ""
    pKal1Tz.SetFocus
    Else
End If
End Sub
Private Sub pKal1tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal1V0.Text = ""
    pKal1V0.SetFocus
    Else
End If
End Sub
Private Sub pKal2X_Click()
pKal2X.Text = ""
End Sub

Private Sub pKal2X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal2Y.Text = ""
    pKal2Y.SetFocus
    Else
End If
End Sub
Private Sub pKal2y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal2h.Text = ""
    pKal2h.SetFocus
    Else
End If
End Sub
Private Sub pKal2h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal2ON.Text = ""
    pKal2ON.SetFocus
    Else
End If
End Sub
Private Sub pKal2on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal2Tz.Text = ""
    pKal2Tz.SetFocus
    Else
End If
End Sub
Private Sub pKal2tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal2V0.Text = ""
    pKal2V0.SetFocus
    Else
End If
End Sub
Private Sub pKal3X_Click()
pKal3X.Text = ""
End Sub

Private Sub pKal3X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal3Y.Text = ""
    pKal3Y.SetFocus
    Else
End If
End Sub
Private Sub pKal3y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal3h.Text = ""
    pKal3h.SetFocus
    Else
End If
End Sub
Private Sub pKal3h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal3ON.Text = ""
    pKal3ON.SetFocus
    Else
End If
End Sub
Private Sub pKal3on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal3Tz.Text = ""
    pKal3Tz.SetFocus
    Else
End If
End Sub
Private Sub pKal3tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pKal3V0.Text = ""
    pKal3V0.SetFocus
    Else
End If
End Sub

Private Sub pOsk1X_Click()
pOsk1X.Text = ""
End Sub

Private Sub pOsk1X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk1Y.Text = ""
    pOsk1Y.SetFocus
    Else
End If
End Sub
Private Sub pOsk1y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk1h.Text = ""
    pOsk1h.SetFocus
    Else
End If
End Sub
Private Sub pOsk1h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk1ON.Text = ""
    pOsk1ON.SetFocus
    Else
End If
End Sub
Private Sub pOsk1on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk1Tz.Text = ""
    pOsk1Tz.SetFocus
    Else
End If
End Sub

Private Sub pOsk1tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk1V0.Text = ""
    pOsk1V0.SetFocus
    Else
End If
End Sub
Private Sub pOsk2X_Click()
pOsk2X.Text = ""
End Sub

Private Sub pOsk2X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk2Y.Text = ""
    pOsk2Y.SetFocus
    Else
End If
End Sub
Private Sub pOsk2y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk2h.Text = ""
    pOsk2h.SetFocus
    Else
End If
End Sub
Private Sub pOsk2h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk2ON.Text = ""
    pOsk2ON.SetFocus
    Else
End If
End Sub
Private Sub pOsk2on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk2Tz.Text = ""
    pOsk2Tz.SetFocus
    Else
End If
End Sub

Private Sub pOsk2tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk2V0.Text = ""
    pOsk2V0.SetFocus
    Else
End If
End Sub

Private Sub pOsk3X_Click()
pOsk3X.Text = ""
End Sub

Private Sub pOsk3X_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk3Y.Text = ""
    pOsk3Y.SetFocus
    Else
End If
End Sub
Private Sub pOsk3y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk3h.Text = ""
    pOsk3h.SetFocus
    Else
End If
End Sub
Private Sub pOsk3h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk3ON.Text = ""
    pOsk3ON.SetFocus
    Else
End If
End Sub
Private Sub pOsk3on_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk3Tz.Text = ""
    pOsk3Tz.SetFocus
    Else
End If
End Sub

Private Sub pOsk3tz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pOsk3V0.Text = ""
    pOsk3V0.SetFocus
    Else
End If
End Sub


