VERSION 5.00
Begin VB.Form VuvodNP 
   BackColor       =   &H0080FF80&
   Caption         =   "¬˚‚Ó‰ Õœ"
   ClientHeight    =   9330
   ClientLeft      =   1125
   ClientTop       =   555
   ClientWidth     =   13935
   BeginProperty Font 
      Name            =   "Cambria"
      Size            =   12
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   13935
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "œÓÍ‡Á‡Ú¸ Õ«Œ"
      Height          =   900
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4000
      Width           =   1700
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "œÓÍ‡Á‡Ú¸ ÷ÂÎË"
      Height          =   900
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2800
      Width           =   1700
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "œÓÍ‡Á‡Ú¸ Œœ"
      Height          =   900
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1600
      Width           =   1700
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "œÓÍ‡Á‡Ú¸ Õœ"
      Height          =   900
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   400
      Width           =   1700
   End
End
Attribute VB_Name = "VuvodNP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Command1_Click()
Cls
Dim t1 As String, t2 As String, t3 As String, t4 As String
941 Open "D:\YO_NA\KNPTABL" For Input As #1
Do While Not EOF(1)
For i = 1 To 28
If EOF(1) Then GoTo 942
 Input #1, t1, t2, t3, t4
Print t1, t2, t3, t4
Next i
Y = MsgBox("œ¿”«¿", vbOKOnly, "œ¿”«¿")
Cls
Loop
942 Close #1
943
End Sub

Private Sub Command2_Click()
Cls
Dim t1 As String, t2 As String, t3 As String, t4 As String
931 Open "D:\YO_NA\OPTABL" For Input As #1
Do While Not EOF(1)
For i = 1 To 28
If EOF(1) Then GoTo 932
 Input #1, t1, t2, t3, t4
Print t1, t2, t3, t4
Next i
Y = MsgBox("œ¿”«¿", vbOKOnly, "œ¿”«¿")
Cls
Loop
932 Close #1
End Sub

Private Sub Command3_Click()
Cls
Dim t1 As String, t2 As String, t3 As String, t4 As String, t5 As String, t6 As String
931 Open "D:\YO_NA\zeli" For Input As #1
Do While Not EOF(1)
For i = 1 To 28
If EOF(1) Then GoTo 932
 Input #1, t1, t2, t3, t4, t5, t6
Print t1, t2, t3, t4, t5, t6
Next i
Y = MsgBox("œ¿”«¿", vbOKOnly, "œ¿”«¿")
Cls
Loop
932 Close #1

End Sub

Private Sub Command4_Click()
Cls
Dim t1 As String, t2 As String, t3 As String, t4 As String, t5 As String, t6 As String
931 Open "D:\YO_NA\nzo" For Input As #1
Do While Not EOF(1)
For i = 1 To 28
If EOF(1) Then GoTo 932
 Input #1, t1, t2, t3, t4, t5, t6
Print t1, t2, t3, t4, t5, t6
Next i
Y = MsgBox("œ¿”«¿", vbOKOnly, "œ¿”«¿")
Cls
Loop
932 Close #1
End Sub
