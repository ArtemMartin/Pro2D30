VERSION 5.00
Begin VB.Form Komandafrm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Команда"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   18960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "ВЫХОД"
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
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6800
      Width           =   4000
   End
   Begin VB.TextBox pvKomandu 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   18375
   End
End
Attribute VB_Name = "Komandafrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Komandafrm

End Sub
