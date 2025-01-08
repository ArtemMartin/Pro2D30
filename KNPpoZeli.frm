VERSION 5.00
Begin VB.Form KNPpoZeli 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ÊÍÏ ïî Öåëè"
   ClientHeight    =   5250
   ClientLeft      =   3120
   ClientTop       =   2955
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   14.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3147.751
   ScaleMode       =   0  'User
   ScaleWidth      =   2485
   Begin VB.TextBox pMc5 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6100
      TabIndex        =   34
      Text            =   "0"
      Top             =   4200
      Width           =   1000
   End
   Begin VB.TextBox pMc4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6100
      TabIndex        =   33
      Text            =   "0"
      Top             =   3200
      Width           =   1000
   End
   Begin VB.TextBox pMc3 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6100
      TabIndex        =   32
      Text            =   "0"
      Top             =   2200
      Width           =   1000
   End
   Begin VB.TextBox pMc2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6100
      TabIndex        =   31
      Text            =   "0"
      Top             =   1200
      Width           =   1000
   End
   Begin VB.TextBox pMc1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6100
      TabIndex        =   30
      Text            =   "0"
      Top             =   200
      Width           =   1000
   End
   Begin VB.TextBox pD5 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4100
      TabIndex        =   24
      Text            =   "0"
      Top             =   4200
      Width           =   1000
   End
   Begin VB.TextBox pD4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4100
      TabIndex        =   23
      Text            =   "0"
      Top             =   3200
      Width           =   1000
   End
   Begin VB.TextBox pD3 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4100
      TabIndex        =   22
      Text            =   "0"
      Top             =   2200
      Width           =   1000
   End
   Begin VB.TextBox pD2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4100
      TabIndex        =   21
      Text            =   "0"
      Top             =   1200
      Width           =   1000
   End
   Begin VB.TextBox pD1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4100
      TabIndex        =   20
      Text            =   "0"
      Top             =   200
      Width           =   1000
   End
   Begin VB.TextBox pA5 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2200
      TabIndex        =   14
      Text            =   "0"
      Top             =   4200
      Width           =   1000
   End
   Begin VB.TextBox pA4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2200
      TabIndex        =   13
      Text            =   "0"
      Top             =   3200
      Width           =   1000
   End
   Begin VB.TextBox pA3 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2200
      TabIndex        =   12
      Text            =   "0"
      Top             =   2200
      Width           =   1000
   End
   Begin VB.TextBox pA2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2200
      TabIndex        =   11
      Text            =   "0"
      Top             =   1200
      Width           =   1000
   End
   Begin VB.TextBox pA1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2200
      TabIndex        =   10
      Text            =   "0"
      Top             =   200
      Width           =   1000
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ìö="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5350
      TabIndex        =   29
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ìö="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5350
      TabIndex        =   28
      Top             =   3200
      Width           =   600
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ìö="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5350
      TabIndex        =   27
      Top             =   2200
      Width           =   600
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ìö="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5350
      TabIndex        =   26
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ìö="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5350
      TabIndex        =   25
      Top             =   200
      Width           =   600
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ä="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3500
      TabIndex        =   19
      Top             =   4200
      Width           =   500
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ä="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3500
      TabIndex        =   18
      Top             =   3200
      Width           =   500
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ä="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3500
      TabIndex        =   17
      Top             =   2200
      Width           =   500
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ä="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3500
      TabIndex        =   16
      Top             =   1200
      Width           =   500
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ä="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3500
      TabIndex        =   15
      Top             =   200
      Width           =   500
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "À="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1600
      TabIndex        =   9
      Top             =   4200
      Width           =   500
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "À="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1600
      TabIndex        =   8
      Top             =   3200
      Width           =   500
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "À="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1600
      TabIndex        =   7
      Top             =   2200
      Width           =   500
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "À="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1600
      TabIndex        =   6
      Top             =   1200
      Width           =   500
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "À="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1600
      TabIndex        =   5
      Top             =   200
      Width           =   500
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ÊÍÏ ¹5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   200
      TabIndex        =   4
      Top             =   4200
      Width           =   1300
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ÊÍÏ ¹4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   200
      TabIndex        =   3
      Top             =   3200
      Width           =   1300
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ÊÍÏ ¹3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   200
      TabIndex        =   2
      Top             =   2200
      Width           =   1300
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ÊÍÏ ¹2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   200
      TabIndex        =   1
      Top             =   1200
      Width           =   1300
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ÊÍÏ ¹1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Width           =   1300
   End
End
Attribute VB_Name = "KNPpoZeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload KNPpoZeli
End Sub
