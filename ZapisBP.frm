VERSION 5.00
Begin VB.Form ZapisBP 
   BackColor       =   &H0080FF80&
   Caption         =   "Запись боевого порядка"
   ClientHeight    =   5010
   ClientLeft      =   3120
   ClientTop       =   3450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   9705
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Записать ОП"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   100
      TabIndex        =   10
      Top             =   2600
      Width           =   9500
      Begin VB.CommandButton Command2 
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
         Height          =   500
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1400
         Width           =   2000
      End
      Begin VB.TextBox phop 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7500
         TabIndex        =   18
         Text            =   "0"
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox pYop 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5400
         TabIndex        =   16
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pXop 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3200
         TabIndex        =   14
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pnOP 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   12
         Text            =   "0"
         Top             =   600
         Width           =   1300
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7100
         TabIndex        =   17
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4900
         TabIndex        =   15
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Х="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2700
         TabIndex        =   13
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ ОП"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Запись НП"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   9500
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
         Height          =   500
         Left            =   6500
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1400
         Width           =   2000
      End
      Begin VB.TextBox phnp 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7500
         TabIndex        =   8
         Text            =   "0"
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox pYnp 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5400
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pXnp 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3200
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox pnNP 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   2
         Text            =   "0"
         Top             =   600
         Width           =   1300
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7100
         TabIndex        =   7
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4900
         TabIndex        =   5
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Х="
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2700
         TabIndex        =   3
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ НП"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   800
      End
   End
End
Attribute VB_Name = "ZapisBP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Nnp As String
Dim X As Single, Y As Single, h As Single
Nnp = pnNP: X = pXnp: Y = pYnp: h = phnp
801 Open "D:\YO_NA\knptabl" For Append As #1
        Write #1, Nnp, X, Y, h
        Close #1
pnNP.Text = 0: pXnp.Text = 0: pYnp.Text = 0: phnp.Text = 0
BP.pplNP1.AddItem (Nnp): BP.pplNP2.AddItem (Nnp): BP.pplNP3.AddItem (Nnp): BP.pplNP4.AddItem (Nnp)
BP.pplNP5.AddItem (Nnp)
End Sub

Private Sub Command2_Click()
Dim Nop As String
Dim X As Single, Y As Single, h As Single
Nop = pnOP: X = pXop: Y = pYop: h = phop
801 Open "D:\YO_NA\optabl" For Append As #1
        Write #1, Nop, X, Y, h
        Close #1
pnOP.Text = 0: pXop.Text = 0: pYop.Text = 0: phop.Text = 0
BP.pplOP1.AddItem (Nop): BP.pplOP2.AddItem (Nop): BP.pplOP3.AddItem (Nop)
End Sub

Private Sub pnNP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pXnp.Text = ""
pXnp.SetFocus
End If
End Sub
Private Sub pXnp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYnp.Text = ""
pYnp.SetFocus
End If
End Sub
Private Sub pYnp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phnp.Text = ""
phnp.SetFocus
End If
End Sub
Private Sub pnOP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pXop.Text = ""
pXop.SetFocus
End If
End Sub
Private Sub pXop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
pYop.Text = ""
pYop.SetFocus
End If
End Sub
Private Sub pYop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
phop.Text = ""
phop.SetFocus
End If
End Sub

