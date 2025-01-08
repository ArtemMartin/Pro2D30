VERSION 5.00
Begin VB.Form RaschSredn 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Расчет среднего"
   ClientHeight    =   6840
   ClientLeft      =   3120
   ClientTop       =   2955
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   6735
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "РАСЧИТАТЬ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3360
      Width           =   6015
   End
   Begin VB.CommandButton otprvReper 
      BackColor       =   &H0000FF00&
      Caption         =   "Отправить в Репер"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   5300
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5000
      Width           =   1300
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Отправить в Сопряженку"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   3200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5000
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Отправить в Х, У"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   1600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5000
      Width           =   1300
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Отправить в ДАК"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   100
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5000
      Width           =   1300
   End
   Begin VB.TextBox pvsr3 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4600
      TabIndex        =   19
      Text            =   "0"
      Top             =   4000
      Width           =   1500
   End
   Begin VB.TextBox pvsr2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2800
      TabIndex        =   18
      Text            =   "0"
      Top             =   4000
      Width           =   1500
   End
   Begin VB.TextBox pvsr1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1000
      TabIndex        =   17
      Text            =   "0"
      Top             =   4000
      Width           =   1500
   End
   Begin VB.TextBox p43 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4600
      TabIndex        =   15
      Text            =   "0"
      Top             =   2700
      Width           =   1500
   End
   Begin VB.TextBox p33 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4600
      TabIndex        =   14
      Text            =   "0"
      Top             =   1900
      Width           =   1500
   End
   Begin VB.TextBox p23 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4600
      TabIndex        =   13
      Text            =   "0"
      Top             =   1100
      Width           =   1500
   End
   Begin VB.TextBox p13 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4600
      TabIndex        =   12
      Text            =   "0"
      Top             =   300
      Width           =   1500
   End
   Begin VB.TextBox p42 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2800
      TabIndex        =   11
      Text            =   "0"
      Top             =   2700
      Width           =   1500
   End
   Begin VB.TextBox p32 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2800
      TabIndex        =   10
      Text            =   "0"
      Top             =   1900
      Width           =   1500
   End
   Begin VB.TextBox p22 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2800
      TabIndex        =   9
      Text            =   "0"
      Top             =   1100
      Width           =   1500
   End
   Begin VB.TextBox p12 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2800
      TabIndex        =   8
      Text            =   "0"
      Top             =   300
      Width           =   1500
   End
   Begin VB.TextBox p41 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1000
      TabIndex        =   7
      Text            =   "0"
      Top             =   2700
      Width           =   1500
   End
   Begin VB.TextBox p31 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1000
      TabIndex        =   6
      Text            =   "0"
      Top             =   1900
      Width           =   1500
   End
   Begin VB.TextBox p21 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1000
      TabIndex        =   5
      Text            =   "0"
      Top             =   1100
      Width           =   1500
   End
   Begin VB.TextBox p11 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1000
      TabIndex        =   4
      Text            =   "0"
      Top             =   300
      Width           =   1500
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ср"
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
      TabIndex        =   16
      Top             =   4000
      Width           =   600
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4="
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
      Top             =   2700
      Width           =   600
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3="
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
      Top             =   1900
      Width           =   600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2="
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
      Top             =   1100
      Width           =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1="
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
      Top             =   300
      Width           =   600
   End
End
Attribute VB_Name = "RaschSredn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ch11 As Single, ch12 As Single, ch13 As Single
Dim ch21 As Single, ch22 As Single, ch23 As Single
Dim ch31 As Single, ch32 As Single, ch33 As Single
Dim ch41 As Single, ch42 As Single, ch43 As Single
Dim chsr1 As Single, chsr2 As Single, chsr3 As Single

Private Sub Command1_Click()
Pristrelka.pAr.Text = chsr1
Pristrelka.pDr.Text = chsr2
Pristrelka.pMr.Text = chsr3
Unload RaschSredn
End Sub

Private Sub Command2_Click()
Pristrelka.pXr.Text = chsr1: Pristrelka.pYr.Text = chsr2
Unload RaschSredn
End Sub

Private Sub Command3_Click()
Pristrelka.pArLev.Text = chsr1: Pristrelka.pArPrav.Text = chsr2
Y = MsgBox("Мц на Левый пункт ?", vbYesNo, "Куда поставить угол места")
If Y = vbYes Then
    Pristrelka.pMrLev.Text = chsr3: Pristrelka.pMrPrav.Text = 0
        Else
            Pristrelka.pMrPrav.Text = chsr3: Pristrelka.pMrLev.Text = 0
End If
Unload RaschSredn
End Sub

Private Sub Command5_Click()
ch11 = p11: ch12 = p12: ch13 = p13
ch21 = p21: ch22 = p22: ch23 = p23
ch31 = p31: ch32 = p32: ch33 = p33
ch41 = p41: ch42 = p42: ch43 = p43
If ch41 <> 0 Then
    chsr1 = Round((ch11 + ch21 + ch31 + ch41) / 4)
    chsr2 = Round((ch12 + ch22 + ch32 + ch42) / 4)
    chsr3 = Round((ch13 + ch23 + ch33 + ch43) / 4)
    ElseIf ch41 = 0 And ch31 <> 0 Then
        chsr1 = Round((ch11 + ch21 + ch31) / 3)
        chsr2 = Round((ch12 + ch22 + ch32) / 3)
        chsr3 = Round((ch13 + ch23 + ch33) / 3)
        ElseIf chs41 = 0 And ch31 = 0 Then
            chsr1 = Round((ch11 + ch21) / 2)
            chsr2 = Round((ch12 + ch22) / 2)
            chsr3 = Round((ch13 + ch23) / 2)
            Else
End If
 pvsr1.Text = chsr1: pvsr2.Text = chsr2: pvsr3.Text = chsr3
End Sub

Private Sub otprvReper_Click()
otpravkavREPER.Show
End Sub
Private Sub p11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p12.Text = ""
p12.SetFocus
End If
End Sub
Private Sub p12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p13.Text = ""
p13.SetFocus
End If
End Sub
Private Sub p13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p21.Text = ""
p21.SetFocus
End If
End Sub
Private Sub p21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p22.Text = ""
p22.SetFocus
End If
End Sub
Private Sub p22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p23.Text = ""
p23.SetFocus
End If
End Sub
Private Sub p23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p31.Text = ""
p31.SetFocus
End If
End Sub
Private Sub p31_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p32.Text = ""
p32.SetFocus
End If
End Sub
Private Sub p32_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p33.Text = ""
p33.SetFocus
End If
End Sub
Private Sub p33_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p41.Text = ""
p41.SetFocus
End If
End Sub
Private Sub p41_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p42.Text = ""
p42.SetFocus
End If
End Sub
Private Sub p42_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
p43.Text = ""
p43.SetFocus
End If
End Sub

