VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PGZ"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton vuxod 
      Caption         =   "Buxod"
      Height          =   735
      Left            =   7440
      TabIndex        =   13
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox phnp 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox pYnp 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox pXnp 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox pHz 
      Height          =   735
      Left            =   1080
      TabIndex        =   6
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox pYz 
      Height          =   855
      Left            =   960
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox pXz 
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox pMz 
      Height          =   855
      Left            =   4800
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox pDz 
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox pAz 
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reshit"
      Height          =   975
      Left            =   7440
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Zel"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "NP"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Zasechka"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Az  As Single
Dim Dz As Single
Dim Mz As Single
Dim Xz As Single
Dim Yz As Single
Dim Hz As Single
Dim Xnp As Single
Dim Ynp As Single
Dim hnp As Single
Private Sub Command1_Click()
Az = pAz: Dz = pDz: Mz = pMz: Xnp = pXnp: Ynp = pYnp: hnp = phnp
Az = Az / 100
Xz = Cos(Az * 6 * 3.141592 / 180) * Dz + Xnp
Yz = Sin(Az * 6 * 3.141592 / 180) * Dz + Ynp
Hz = Mz * (Dz / 1000) * 1.05 + hnp
pXz.Text = Format(Xz, "0.0")
pYz.Text = Format(Yz, "0.0")
pHz.Text = Format(Hz, "0.0")
End Sub

Private Sub vuxod_Click()
End
End Sub
