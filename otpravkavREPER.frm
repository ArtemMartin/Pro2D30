VERSION 5.00
Begin VB.Form otpravkavREPER 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Отправка в РЕПЕР"
   ClientHeight    =   2070
   ClientLeft      =   3120
   ClientTop       =   3450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   7200
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Х, У"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   500
      Width           =   2000
   End
   Begin VB.CommandButton otprSopr 
      BackColor       =   &H0080FF80&
      Caption         =   "Сопряженка"
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
      Left            =   2500
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   500
      Width           =   2000
   End
   Begin VB.CommandButton otprvDAK 
      BackColor       =   &H0080FF80&
      Caption         =   "ДАК"
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
      Left            =   200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   500
      Width           =   2000
   End
End
Attribute VB_Name = "otpravkavREPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
REPER.pXre.Text = RaschSredn.pvsr1: REPER.pYre.Text = RaschSredn.pvsr2: REPER.phre.Text = RaschSredn.pvsr3
Unload otpravkavREPER
Unload RaschSredn
REPER.Show
End Sub

Private Sub otprSopr_Click()
Y = MsgBox("Мц на левый пункт?", vbYesNo, "Мц")
If Y = vbYes Then
    REPER.pMl.Text = RaschSredn.pvsr3
    Else
        REPER.pMp.Text = RaschSredn.pvsr3
End If
REPER.pAl.Text = RaschSredn.pvsr1: REPER.pAp.Text = RaschSredn.pvsr2
Unload otpravkavREPER
Unload RaschSredn
REPER.Show
End Sub

Private Sub otprvDAK_Click()
REPER.pAre.Text = RaschSredn.pvsr1
REPER.pDre.Text = RaschSredn.pvsr2
REPER.pMre.Text = RaschSredn.pvsr3
Unload otpravkavREPER
Unload RaschSredn
REPER.Show
End Sub
