VERSION 5.00
Begin VB.Form Razrabotchik 
   BackColor       =   &H0080FFFF&
   Caption         =   "Разработчик"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Закрыть"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   7000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Разработал А.В. Ступак Донецк 2016-2019. Предназначена для управления огнем дивизиона Д30. В помощь и во славу Отечества!"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1500
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   8400
   End
End
Attribute VB_Name = "Razrabotchik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Razrabotchik.Hide
End Sub
