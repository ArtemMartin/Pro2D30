VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell App.Path & "\sas\sasplanet", vbNormalFocus
End Sub
