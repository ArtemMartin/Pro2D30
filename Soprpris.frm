VERSION 5.00
Begin VB.Form Soprpris 
   Caption         =   "Пристрелка Сопряженным"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   450
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
   ScaleHeight     =   8505
   ScaleWidth      =   7770
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Решить"
      Height          =   1000
      Left            =   6250
      TabIndex        =   19
      Top             =   2100
      Width           =   1300
   End
   Begin VB.TextBox Text7 
      Height          =   500
      Left            =   2700
      TabIndex        =   18
      Text            =   "0"
      Top             =   7300
      Width           =   1700
   End
   Begin VB.TextBox Text6 
      Height          =   500
      Left            =   2700
      TabIndex        =   17
      Text            =   "0"
      Top             =   6300
      Width           =   1700
   End
   Begin VB.TextBox Text5 
      Height          =   500
      Left            =   2700
      TabIndex        =   16
      Text            =   "0"
      Top             =   5300
      Width           =   1700
   End
   Begin VB.TextBox Text4 
      Height          =   500
      Left            =   4250
      TabIndex        =   15
      Text            =   "0"
      Top             =   3100
      Width           =   1700
   End
   Begin VB.TextBox Text3 
      Height          =   500
      Left            =   4250
      TabIndex        =   14
      Text            =   "0"
      Top             =   2100
      Width           =   1700
   End
   Begin VB.TextBox Text2 
      Height          =   500
      Left            =   1050
      TabIndex        =   13
      Text            =   "0"
      Top             =   3100
      Width           =   1700
   End
   Begin VB.TextBox Text1 
      Height          =   500
      Left            =   1050
      TabIndex        =   12
      Text            =   "0"
      Top             =   2100
      Width           =   1700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Выход"
      Height          =   2175
      Left            =   9840
      TabIndex        =   0
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label11 
      Caption         =   "Ур="
      Height          =   500
      Left            =   2000
      TabIndex        =   11
      Top             =   7300
      Width           =   500
   End
   Begin VB.Label Label10 
      Caption         =   "Дов="
      Height          =   500
      Left            =   2000
      TabIndex        =   10
      Top             =   6300
      Width           =   700
   End
   Begin VB.Label Label9 
      Caption         =   "Д="
      Height          =   500
      Left            =   2000
      TabIndex        =   9
      Top             =   5300
      Width           =   400
   End
   Begin VB.Label Label8 
      Caption         =   "     Корректура"
      Height          =   500
      Left            =   2000
      TabIndex        =   8
      Top             =   4300
      Width           =   2400
   End
   Begin VB.Label Label7 
      Caption         =   "Мцр="
      Height          =   500
      Left            =   3500
      TabIndex        =   7
      Top             =   3100
      Width           =   700
   End
   Begin VB.Label Label6 
      Caption         =   "Ар="
      Height          =   500
      Left            =   3500
      TabIndex        =   6
      Top             =   2100
      Width           =   500
   End
   Begin VB.Label Label5 
      Caption         =   "Мцр="
      Height          =   500
      Left            =   300
      TabIndex        =   5
      Top             =   3100
      Width           =   700
   End
   Begin VB.Label Label4 
      Caption         =   "Ар="
      Height          =   500
      Left            =   300
      TabIndex        =   4
      Top             =   2100
      Width           =   500
   End
   Begin VB.Label Label3 
      Caption         =   "         Правый"
      Height          =   500
      Left            =   3500
      TabIndex        =   3
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label Label2 
      Caption         =   "          Левый"
      Height          =   500
      Left            =   300
      TabIndex        =   2
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "             Пристрелка сопряженным"
      Height          =   500
      Left            =   300
      TabIndex        =   1
      Top             =   360
      Width           =   5700
   End
End
Attribute VB_Name = "Soprpris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Soprpris
End Sub

