VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10995
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl Zvyk 
      Height          =   1935
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3413
      _Version        =   393216
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   "E:\Proekt2ñ1\Kalimba.mp3"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MMControl1_Done(NotifyCode As Integer)
Zvyk.DeviceType = "WaveAudio"
Zvyk.FileName = App.Path & "\kalimba.mp3"
Zvyk.Command = "Open"
Zvyk.Command = "Play"

End Sub

Private Sub Zvyk_Done(NotifyCode As Integer)
Zvyk.DeviceType = "WaveAudio"
Zvyk.FileName = App.Path & "\kalimba.mp3"
Zvyk.Command = "Open"
Zvyk.Command = "Play"
End Sub
