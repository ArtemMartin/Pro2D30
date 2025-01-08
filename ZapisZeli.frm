VERSION 5.00
Begin VB.Form ZapisZeli 
   BackColor       =   &H0080FF80&
   Caption         =   "Добавление в архив"
   ClientHeight    =   5520
   ClientLeft      =   1125
   ClientTop       =   2445
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   14250
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Записать НЗО"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2600
      Left            =   100
      TabIndex        =   14
      Top             =   2800
      Width           =   14000
      Begin VB.CommandButton zapNZO 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Записать"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1400
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1500
         Width           =   2200
      End
      Begin VB.TextBox phc 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   12500
         TabIndex        =   26
         Text            =   "0"
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox pYp 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   10300
         TabIndex        =   24
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pXp 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   8100
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pYl 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5900
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pXl 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3600
         TabIndex        =   18
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pnNZO 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1400
         TabIndex        =   16
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "hc="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   11900
         TabIndex        =   25
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Уп="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   9700
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Хп="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   7500
         TabIndex        =   21
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ул="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5300
         TabIndex        =   19
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Хл="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3000
         TabIndex        =   17
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Назв. ЗО"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   100
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame ZapisZeli 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Записать цель"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2600
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   14000
      Begin VB.TextBox Pgl 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   12000
         TabIndex        =   13
         Text            =   "0"
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox pFr 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   10000
         TabIndex        =   11
         Text            =   "0"
         Top             =   600
         Width           =   1000
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Записать"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1400
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1500
         Width           =   2200
      End
      Begin VB.TextBox phz 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   8100
         TabIndex        =   8
         Text            =   "0"
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox pYz 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5900
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pXz 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3600
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pnZeli 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1400
         TabIndex        =   2
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Гл="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   11200
         TabIndex        =   12
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фр="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   9300
         TabIndex        =   10
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "hц="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7500
         TabIndex        =   7
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Уц="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5300
         TabIndex        =   5
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Хц="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Цели"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   200
         TabIndex        =   1
         Top             =   600
         Width           =   1200
      End
   End
End
Attribute VB_Name = "ZapisZeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xar As String
Dim X As Single, Y As Single, h As Single, Fr As Single, Gl As Single
Xar = pnZeli: X = pXz: Y = pYz: h = phz: Fr = pFr: Gl = Pgl
801 Open "D:\YO_NA\Zeli" For Append As #1
        Write #1, Xar, X, Y, h, Fr, Gl
        Close #1
pnZeli.Text = 0: pXz.Text = 0: pYz.Text = 0: phz.Text = 0: Pgl.Text = 0: pFr.Text = 0
OZ.pplZel.AddItem (Xar)
End Sub

Private Sub pnZeli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pXz.Text = ""
pXz.SetFocus
End If
End Sub
Private Sub pXz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYz.Text = ""
pYz.SetFocus
End If
End Sub
Private Sub pYz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phz.Text = ""
phz.SetFocus
End If
End Sub
Private Sub phz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pFr.Text = ""
pFr.SetFocus
End If
End Sub
Private Sub pFr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Pgl.Text = ""
Pgl.SetFocus
End If
End Sub
Private Sub pnNZO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pXl.Text = ""
pXl.SetFocus
End If
End Sub
Private Sub pXl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYl.Text = ""
pYl.SetFocus
End If
End Sub
Private Sub pYl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pXp.Text = ""
pXp.SetFocus
End If
End Sub
Private Sub pXp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYp.Text = ""
pYp.SetFocus
End If
End Sub
Private Sub pYp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phc.Text = ""
phc.SetFocus
End If
End Sub

Private Sub zapNZO_Click()
Dim nnZo As String
Dim Xl As Single, Yl As Single, hc As Single, Xp As Single, Yp As Single
nnZo = pnNZO: Xl = pXl: Yl = pYl: Xp = pXp: Yp = pYp: hc = phc
801 Open "D:\YO_NA\nzo" For Append As #1
        Write #1, nnZo, Xl, Yl, Xp, Yp, hc
        Close #1
pnNZO.Text = 0: pXl.Text = 0: pYl.Text = 0: pXp.Text = 0: pYp.Text = 0: phc.Text = 0
End Sub
