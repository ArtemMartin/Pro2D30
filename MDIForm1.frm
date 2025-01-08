VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.MDIForm SYO 
   BackColor       =   &H0000C0C0&
   Caption         =   "Управление огнем"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14940
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton Pristrelkakn 
         BackColor       =   &H0000FFFF&
         Caption         =   "Пристрелка"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1500
      End
   End
End
Attribute VB_Name = "SYO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public A As Single
Public D As Single
Public Mz As Single
Public Xnp As Single
Public Ynp As Single
Public hnp As Single
Public Xz As Single
Public Yz As Single
Public hz As Single
Public Xop As Single
Public Yop As Single
Public hop As Single
Public Dt As Single
Public Yr As Single
Public dh, Pi, Ygolt, OH, dovort As Single



Private Sub Command1_Click()
Soprpris.Show
End Sub

Private Sub Pristrelkakn_Click()
Pristrelka.Show
End Sub
