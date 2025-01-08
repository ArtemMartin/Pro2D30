VERSION 5.00
Begin VB.Form Komanda 
   BackColor       =   &H00808080&
   Caption         =   "Команда"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18225
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6.75
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   18225
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton sendGdynu 
      BackColor       =   &H000080FF&
      Caption         =   "Отправить в ""Ждуны"""
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7600
      Width           =   5000
   End
   Begin VB.CommandButton clickVuxod 
      BackColor       =   &H008080FF&
      Caption         =   "ВЫХОД"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7600
      Width           =   5000
   End
   Begin VB.TextBox pvKomand 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7000
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Komanda.frx":0000
      Top             =   300
      Width           =   17625
   End
End
Attribute VB_Name = "Komanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clickVuxod_Click()
Unload Komanda
End Sub

Private Sub sendGdynu_Click()
'работа с ботом transmitter101
Dim oHttp As Object
Dim sURI As String
Dim str As String
Dim mas() As String
'Dnepr Transmitter
token = "6631245904:AAHO9mLro2qJ9RlKz6KBM52DlS1txdw4NiM"
'Ждуны
chat_id = "-1001757624806"

'Transmitter102
'token = "6240157156:AAHHqUAhr-C-MAPORkJ1CUIBR8Dl9_Re2vE"
'Accept Message
'chat_id = "-1001926672455"

'Нужно использовать символ переноса строки %0A
str = pvKomand
str = RussianStringToURLEncode_New(str)
'чтоб не терять плюс в сообщении
mas = Split(str, "+")
str = Join(mas, "%2B")
'чтоб была пустая строка между ОП
str = polychitStrokySPerenosami(str)

sURI = "https://api.telegram.org/bot" & token & "/sendmessage?chat_id=" & chat_id & "&text=" & str
 
MsgBox sURI, vbInformation, "запрос"
On Error Resume Next
Set oHttp = CreateObject("MSXML2.XMLHTTP")
If Err.Number <> 0 Then
Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
End If
On Error GoTo 0
If oHttp Is Nothing Then Exit Sub
oHttp.Open "GET", sURI, False
oHttp.send
MsgBox oHttp.ResponseText, vbInformation, "ответ"
Set oHttp = Nothing

End Sub
'для русского текста
Function RussianStringToURLEncode_New(ByVal txt As String) As String
    For i = 1 To Len(txt)
        l = Mid(txt, i, 1)
        Select Case AscW(l)
            Case Is > 4095: t = "%" & Hex(AscW(l) \ 64 \ 64 + 224) & "%" & Hex(AscW(l) \ 64) & "%" & Hex(8 * 16 + AscW(l) Mod 64)
            Case Is > 127: t = "%" & Hex(AscW(l) \ 64 + 192) & "%" & Hex(8 * 16 + AscW(l) Mod 64)
            Case 32: t = "%20"
            Case Else: t = l
        End Select
        RussianStringToURLEncode_New = RussianStringToURLEncode_New & t
    Next
End Function
Function polychitStrokySPerenosami(ByVal str As String) As String
Dim mas() As String

mas = Split(str, "__")
str = ""
For i = 0 To UBound(mas)
    str = str & mas(i) & "%0A"
    str = str & "%0A"
Next i
polychitStrokySPerenosami = str
End Function
