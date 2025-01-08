VERSION 5.00
Begin VB.Form zasGryp 
   BackColor       =   &H0080FF80&
   Caption         =   "Засечка груповой Цели"
   ClientHeight    =   9825
   ClientLeft      =   3120
   ClientTop       =   3450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleLeft       =   3000
   ScaleMode       =   0  'User
   ScaleTop        =   2000
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "А, Д"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   10935
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Выход"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   2700
         Width           =   1500
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   2175
         Left            =   7700
         TabIndex        =   59
         Top             =   7000
         Width           =   1575
         Begin VB.OptionButton plevkrai 
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   1000
            TabIndex        =   62
            Top             =   1600
            Width           =   375
         End
         Begin VB.OptionButton pzentr 
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   1000
            TabIndex        =   61
            Top             =   1000
            Width           =   495
         End
         Begin VB.OptionButton pprkrai 
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   1000
            TabIndex        =   60
            Top             =   400
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.Label Label37 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Левый край"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9.75
               Charset         =   204
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   100
            TabIndex        =   65
            Top             =   1600
            Width           =   855
         End
         Begin VB.Label Label36 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Центр"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9.75
               Charset         =   204
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   100
            TabIndex        =   64
            Top             =   1000
            Width           =   855
         End
         Begin VB.Label Label35 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Правый край"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9.75
               Charset         =   204
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   100
            TabIndex        =   63
            Top             =   400
            Width           =   855
         End
      End
      Begin VB.OptionButton pkagdomy 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6200
         TabIndex        =   56
         Top             =   8500
         Width           =   495
      End
      Begin VB.OptionButton pvsem 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6200
         TabIndex        =   54
         Top             =   7500
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.CheckBox p3bat 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5000
         TabIndex        =   52
         Top             =   8000
         Value           =   1  'Checked
         Width           =   400
      End
      Begin VB.CheckBox p2bat 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5000
         TabIndex        =   51
         Top             =   7500
         Value           =   1  'Checked
         Width           =   400
      End
      Begin VB.CheckBox p1bat 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5000
         TabIndex        =   50
         Top             =   7000
         Value           =   1  'Checked
         Width           =   400
      End
      Begin VB.TextBox pGlpor 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2500
         TabIndex        =   45
         Text            =   "0"
         Top             =   7500
         Width           =   1000
      End
      Begin VB.TextBox pFrpor 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   700
         TabIndex        =   43
         Text            =   "0"
         Top             =   7500
         Width           =   1000
      End
      Begin VB.TextBox pvhzc 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9200
         TabIndex        =   39
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvYzc 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   7000
         TabIndex        =   37
         Text            =   "0"
         Top             =   5200
         Width           =   1500
      End
      Begin VB.TextBox pvXzc 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4700
         TabIndex        =   35
         Text            =   "0"
         Top             =   5200
         Width           =   1500
      End
      Begin VB.ComboBox pNpodr 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "ZasGrypZeli.frx":0000
         Left            =   9400
         List            =   "ZasGrypZeli.frx":000D
         TabIndex        =   32
         Text            =   "1"
         Top             =   500
         Width           =   1000
      End
      Begin VB.TextBox pvGl 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2700
         TabIndex        =   29
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pvFr 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   900
         TabIndex        =   27
         Text            =   "0"
         Top             =   5200
         Width           =   1000
      End
      Begin VB.TextBox pBA 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   25
         Text            =   "0"
         Top             =   4000
         Width           =   1000
      End
      Begin VB.TextBox pDA 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   23
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pPD 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7700
         TabIndex        =   21
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pLD 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2300
         TabIndex        =   19
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.CommandButton reshZasGr 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Решить"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1500
         Width           =   1500
      End
      Begin VB.TextBox pBMc 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6000
         TabIndex        =   16
         Text            =   "0"
         Top             =   4000
         Width           =   1000
      End
      Begin VB.TextBox pPA 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6000
         TabIndex        =   15
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pBD 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4100
         TabIndex        =   12
         Text            =   "0"
         Top             =   4000
         Width           =   1000
      End
      Begin VB.TextBox pLA 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   600
         TabIndex        =   8
         Text            =   "0"
         Top             =   2400
         Width           =   1000
      End
      Begin VB.TextBox pDMc 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6000
         TabIndex        =   3
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox pDD 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4100
         TabIndex        =   2
         Text            =   "0"
         Top             =   1000
         Width           =   1000
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Отмерять от"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7500
         TabIndex        =   58
         Top             =   6500
         Width           =   2000
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Цель"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6000
         TabIndex        =   57
         Top             =   6500
         Width           =   800
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Каждому"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5700
         TabIndex        =   55
         Top             =   8000
         Width           =   1500
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Всем"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6000
         TabIndex        =   53
         Top             =   7000
         Width           =   975
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4000
         TabIndex        =   49
         Top             =   8000
         Width           =   1000
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4000
         TabIndex        =   48
         Top             =   7500
         Width           =   1000
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 Бат"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4000
         TabIndex        =   47
         Top             =   7000
         Width           =   1000
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Привлечь"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3950
         TabIndex        =   46
         Top             =   6500
         Width           =   1500
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Гл="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1920
         TabIndex        =   44
         Top             =   7500
         Width           =   600
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фр="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   100
         TabIndex        =   42
         Top             =   7500
         Width           =   600
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Выбрать поражаемый размер"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   100
         TabIndex        =   41
         Top             =   6500
         Width           =   3400
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Распределение по участкам"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2500
         TabIndex        =   40
         Top             =   6000
         Width           =   4500
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "h="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   8700
         TabIndex        =   38
         Top             =   5200
         Width           =   500
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "У="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   6500
         TabIndex        =   36
         Top             =   5200
         Width           =   615
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Х="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4200
         TabIndex        =   34
         Top             =   5200
         Width           =   495
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "По центру Цели"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5500
         TabIndex        =   33
         Top             =   4680
         Width           =   3255
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "№ Подручной"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   7600
         TabIndex        =   31
         Top             =   500
         Width           =   1700
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Размеры цели"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1000
         TabIndex        =   30
         Top             =   4680
         Width           =   2000
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Гл="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2000
         TabIndex        =   28
         Top             =   5200
         Width           =   700
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Фр="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   100
         TabIndex        =   26
         Top             =   5200
         Width           =   700
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1900
         TabIndex        =   24
         Top             =   4000
         Width           =   500
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1900
         TabIndex        =   22
         Top             =   1000
         Width           =   500
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Д="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7200
         TabIndex        =   20
         Top             =   2400
         Width           =   500
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Д="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1800
         TabIndex        =   18
         Top             =   2400
         Width           =   500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5500
         TabIndex        =   14
         Top             =   2400
         Width           =   500
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Правая граница"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   6200
         TabIndex        =   13
         Top             =   1800
         Width           =   2300
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5400
         TabIndex        =   11
         Top             =   4000
         Width           =   600
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Д="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3600
         TabIndex        =   10
         Top             =   4000
         Width           =   500
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "       Ближняя граница"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2800
         TabIndex        =   9
         Top             =   3400
         Width           =   3400
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "А="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   7
         Top             =   2400
         Width           =   500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Левая граница"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   800
         TabIndex        =   6
         Top             =   1800
         Width           =   2100
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "       Дальняя граница"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2800
         TabIndex        =   5
         Top             =   400
         Width           =   3400
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Мц="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5400
         TabIndex        =   4
         Top             =   1000
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Д="
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3600
         TabIndex        =   1
         Top             =   1000
         Width           =   500
      End
   End
End
Attribute VB_Name = "zasGryp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
zasGryp.Hide
End Sub

Private Sub reshZasGr_Click()
Dim pravA As Single, pravD As Single, levA As Single, levD As Single, dalnA As Single, dalnD As Single, bligA As Single, bligD As Single, md As Single, mb As Single, Mc As Single
Dim Xkp As Single, Ykp As Single, Xop As Single, Yop As Single, Frpor As Single, Glpor As Single
pravA = pPA: pravD = pPD: levA = pLA: levD = pLD: dalnA = pDA: dalnD = pDD: bligA = pBA: bligD = pBD
md = pDMc: mb = pBMc
Mc = (md - mb) / 2 + mb
If OZ.pnkpA = 1 Then
        Xkp = BP.pXkp1: Ykp = BP.pYkp1: hkp = BP.phkp1
        ElseIf OZ.pnkpA = 2 Then
            Xkp = BP.pXkp2: Ykp = BP.pYkp2: hkp = BP.phkp2
         ElseIf OZ.pnkpA = 3 Then
            Xkp = BP.pXkp3: Ykp = BP.pYkp3: hkp = BP.phkp3
               ElseIf OZ.pnkpA = 4 Then
                  Xkp = BP.pXkp4: Ykp = BP.pYkp4: hkp = BP.phkp4
                   ElseIf OZ.pnkpA = 5 Then
                      Xkp = BP.pXkp5: Ykp = BP.pYkp5: hkp = BP.phkp5
    Else
End If
hzc = Round(Mc * ((((dalnD - bligD + 0.001) / 2) + bligD) / 1000) * 1.05 + hkp)
   Pi = 3.14159265358
Xpr = Cos(pravA / 100 * 6 * Pi / 180) * pravD + Xkp
Ypr = Sin(pravA / 100 * 6 * Pi / 180) * pravD + Ykp
Xlev = Cos(levA / 100 * 6 * Pi / 180) * levD + Xkp
Ylev = Sin(levA / 100 * 6 * Pi / 180) * levD + Ykp
Xdaln = Cos(dalnA / 100 * 6 * Pi / 180) * dalnD + Xkp
Ydaln = Sin(dalnA / 100 * 6 * Pi / 180) * dalnD + Ykp
Xblig = Cos(bligA / 100 * 6 * Pi / 180) * bligD + Xkp
Yblig = Sin(bligA / 100 * 6 * Pi / 180) * bligD + Ykp
If pNpodr = 1 Then
    Xop = BP.pX1: Yop = BP.pY1
        ElseIf pNpodr = 2 Then
            Xop = BP.pX2: Yop = BP.pY2
            ElseIf pNpodr = 3 Then
                Xop = BP.pX3: Yop = BP.pY3
    Else
End If
Xc = Xpr: Yc = Ypr
IZMERENIE Xc, Yc, Xop, Yop, Dt, Ygolt
Dt1 = Dt: Ygolt1 = Ygolt
Xc = Xlev: Yc = Ylev
IZMERENIE Xc, Yc, Xop, Yop, Dt, Ygolt
Dt2 = Dt: Ygolt2 = Ygolt
Xc = Xdaln: Yc = Ydaln
IZMERENIE Xc, Yc, Xop, Yop, Dt, Ygolt
Dt3 = Dt: Ygolt3 = Ygolt
Xc = Xblig: Yc = Yblig
IZMERENIE Xc, Yc, Xop, Yop, Dt, Ygolt
Dt4 = Dt: Ygolt4 = Ygolt
If Dt1 < Dt2 And Dt1 < Dt3 And Dt1 < Dt4 Then
    bligD = Dt1
    ElseIf Dt2 < Dt1 And Dt2 < Dt3 And Dt2 < Dt4 Then
        bligD = Dt2
                ElseIf Dt3 < Dt1 And Dt3 < Dt2 And Dt3 < Dt4 Then
                      bligD = Dt3
                            ElseIf Dt4 < Dt1 And Dt4 < Dt3 And Dt4 < Dt2 Then
                                  bligD = Dt4
    Else
End If
If Dt1 > Dt2 And Dt1 > Dt3 And Dt1 < Dt4 Then
    dalnD = Dt1
        ElseIf Dt2 > Dt1 And Dt2 > Dt3 And Dt2 < Dt4 Then
             dalnD = Dt2
             ElseIf Dt3 > Dt2 And Dt3 > Dt1 And Dt3 > Dt4 Then
                dalnD = Dt3
                ElseIf Dt4 > Dt2 And Dt4 > Dt3 And Dt4 < Dt1 Then
                    dalnD = Dt4
    Else
End If
If Ygolt1 > 0 And Ygolt1 < 1500 And Ygolt1 < Ygolt2 And Ygolt2 > 4500 And Ygolt1 < Ygolt3 And Ygolt3 > 4500 And Ygolt1 < Ygolt4 And Ygolt4 > 4500 Then
    Ygolt1 = Ygolt1 + 6000
    ElseIf Ygolt2 > 0 And Ygolt2 < 1500 And Ygolt2 < Ygolt1 And Ygolt1 > 4500 And Ygolt2 < Ygolt3 And Ygolt3 > 4500 And Ygolt2 < Ygolt4 And Ygolt4 > 4500 Then
        Ygolt2 = Ygolt2 + 6000
                ElseIf Ygolt3 > 0 And Ygolt3 < 1500 And Ygolt3 < Ygolt1 And Ygolt1 > 4500 And Ygolt3 < Ygolt2 And Ygolt3 > 4500 And Ygolt3 < Ygolt4 And Ygolt4 > 4500 Then
                    Ygolt3 = Ygolt3 + 6000
                            ElseIf Ygolt4 > 0 And Ygolt4 < 1500 And Ygolt4 < Ygolt1 And Ygolt1 > 4500 And Ygolt4 < Ygolt3 And Ygolt3 > 4500 And Ygolt4 < Ygolt2 And Ygolt2 > 4500 Then
                                Ygolt4 = Ygolt4 + 6000
    Else
End If
If Ygolt1 > Ygolt2 And Ygolt1 > Ygolt3 And Ygolt1 > Ygolt4 Then
    pravA = Ygolt1
        ElseIf Ygolt2 > Ygolt1 And Ygolt2 > Ygolt3 And Ygolt2 > Ygolt4 Then
            pravA = Ygolt2
          ElseIf Ygolt3 > Ygolt1 And Ygolt3 > Ygolt2 And Ygolt3 > Ygolt4 Then
                pravA = Ygolt3
                ElseIf Ygolt4 > Ygolt1 And Ygolt4 > Ygolt3 And Ygolt4 > Ygolt2 Then
                     pravA = Ygolt4
    Else
End If
If Ygolt1 < Ygolt2 And Ygolt1 < Ygolt3 And Ygolt1 < Ygolt4 Then
    levA = Ygolt1
    ElseIf Ygolt2 < Ygolt1 And Ygolt2 < Ygolt3 And Ygolt2 < Ygolt4 Then
           levA = Ygolt2
                ElseIf Ygolt3 < Ygolt1 And Ygolt3 < Ygolt2 And Ygolt3 < Ygolt4 Then
                   levA = Ygolt3
                       ElseIf Ygolt4 < Ygolt1 And Ygolt4 < Ygolt3 And Ygolt4 < Ygolt2 Then
                             levA = Ygolt4
    Else
End If
If pravA > 0 And pravA < 1500 And pravA < levA And levA > 4500 Then
    raznygl = pravA + 6000 - levA
    Else
        raznygl = pravA - levA
End If
        Fr = Round(((raznygl) * (((dalnD - bligD) / 2 + bligD) / 1000) * 1.05) + 0.001)
        Gl = dalnD - bligD
        Azc = Round((raznygl) / 2 + levA)
        If Azc > 6000 Then
                Azc = Azc - 6000
                Else
        End If
        Dzc = Round((dalnD - bligD) / 2 + bligD)
Xzc = Round(Cos((Azc + 0.001) / 100 * 6 * Pi / 180) * Dzc + Xop)
Yzc = Round(Sin((Azc + 0.001) / 100 * 6 * Pi / 180) * Dzc + Yop)
pvFr = Fr: pvGl = Gl: pvXzc = Xzc: pvYzc = Yzc: pvhzc = hzc
Xc = Xzc: Yc = Yzc
IZMERENIE Xc, Yc, Xop, Yop, Dt, Ygolt
Dtzc = Dt: Ygoltzc = Ygolt
Dbliggran = Dtzc - ((Gl + 0.001) / 2)
Xbliggran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * Dbliggran + Xop)
Ybliggran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * Dbliggran + Yop)
Frpor = pFrpor: Glpor = pGlpor
Xprbliggran = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * ((Fr + 0.001) / 2) + Xbliggran)
Yprbliggran = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * ((Fr + 0.001) / 2) + Ybliggran)
If Ygoltzc - 1500 < 0 Then
    Ygoltzc1 = Ygoltzc + 6000
    Else
    Ygoltzc1 = Ygoltzc
Xlevbliggran = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * ((Fr + 0.001) / 2) + Xbliggran)
Ylevbliggran = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * ((Fr + 0.001) / 2) + Ybliggran)
End If
If pvsem = True Then
    If pprkrai = True Then
        Xprsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xprbliggran)
        Yprsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Yprbliggran)
        Xzporc = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * ((Frpor + 0.001) / 2) + Xprsrgran)
        Yzporc = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * ((Frpor + 0.001) / 2) + Yprsrgran)
        OZ.pXc = Xzporc: OZ.pYc = Yzporc: OZ.phc = hzc: OZ.pFrontc = Frpor: OZ.pGlybinac = Glpor
        ElseIf pzentr = True Then
            Xzporc = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * (((Glpor + 0.001) / 2 + Dbliggran)) + Xop)
            Yzporc = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * (((Glpor + 0.001) / 2 + Dbliggran)) + Yop)
            OZ.pXc = Xzporc: OZ.pYc = Yzporc: OZ.phc = hzc: OZ.pFrontc = Frpor: OZ.pGlybinac = Glpor
            ElseIf plevkrai = True Then
                         Xlevsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xlevbliggran)
                        Ylevsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ylevbliggran)
                        Xzporc = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * ((Frpor + 0.001) / 2) + Xlevsrgran)
                        Yzporc = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * ((Frpor + 0.001) / 2) + Ylevsrgran)
                        OZ.pXc = Xzporc: OZ.pYc = Yzporc: OZ.phc = hzc: OZ.pFrontc = Frpor: OZ.pGlybinac = Glpor
        Else
    End If
    ElseIf pkagdomy = True Then
        If p1bat = 1 And p2bat = 1 And p3bat = 1 Then
            doliabatych = Frpor / 6
                    If pprkrai = True Then
                         Xprsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xprbliggran)
                        Yprsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Yprbliggran)
                        X1b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Xprsrgran)
                        Y1b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Yprsrgran)
                        X2b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Xprsrgran)
                        Y2b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Yprsrgran)
                        X3b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 5) + Xprsrgran)
                        Y3b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 5) + Yprsrgran)
                        OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                        OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                        OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                        OZzelkagdform.Show
                        ElseIf pzentr = True Then
                            X2b = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xbliggran)
                            Y2b = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ybliggran)
                            X1b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 2) + X2b)
                            Y1b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 2) + Y2b)
                            X3b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 2) + X2b)
                            Y3b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 2) + Y2b)
                            OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                            OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                            OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                            OZzelkagdform.Show
                            ElseIf plevkrai = True Then
                                 Xlevsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xlevbliggran)
                                Ylevsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ylevbliggran)
                                X3b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Xlevsrgran)
                                Y3b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Ylevsrgran)
                                X2b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Xlevsrgran)
                                Y2b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Ylevsrgran)
                                X1b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 5) + Xlevsrgran)
                                Y1b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 5) + Ylevsrgran)
                                OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                                OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                                OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                                OZzelkagdform.Show
                        Else
                    End If
                ElseIf p1bat = 1 And p2bat = 1 And p3bat = 0 Then
                    doliabatych = Frpor / 4
                                                If pprkrai = True Then
                                                    Xprsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xprbliggran)
                                                    Yprsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Yprbliggran)
                                                    X1b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Xprsrgran)
                                                    Y1b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Yprsrgran)
                                                    X2b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Xprsrgran)
                                                    Y2b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Yprsrgran)
                                                    OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                                                    OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                                                    OZzelkagdform.pXc3 = 0: OZzelkagdform.pYc3 = 0: OZzelkagdform.phc3 = 0: OZzelkagdform.pFrc3 = 0: OZzelkagdform.pGlc3 = 0
                                                    OZzelkagdform.Show
                                                    ElseIf pzentr = True Then
                                                            Xb = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xbliggran)
                                                            Yb = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ybliggran)
                                                            X1b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Xb)
                                                            Y1b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Yb)
                                                            X2b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Xb)
                                                            Y2b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Yb)
                                                            OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                                                            OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                                                            OZzelkagdform.pXc3 = 0: OZzelkagdform.pYc3 = 0: OZzelkagdform.phc3 = 0: OZzelkagdform.pFrc3 = 0: OZzelkagdform.pGlc3 = 0
                                                            OZzelkagdform.Show
                                                            ElseIf plevkrai = True Then
                                                                Xlevsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xlevbliggran)
                                                                Ylevsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ylevbliggran)
                                                                X2b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Xlevsrgran)
                                                                Y2b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Ylevsrgran)
                                                                X1b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Xlevsrgran)
                                                                Y1b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Ylevsrgran)
                                                                OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                                                                OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                                                                  OZzelkagdform.pXc3 = 0: OZzelkagdform.pYc3 = 0: OZzelkagdform.phc3 = 0: OZzelkagdform.pFrc3 = 0: OZzelkagdform.pGlc3 = 0
                                                                OZzelkagdform.Show

                                                        Else
                                                End If
                                       ElseIf p1bat = 1 And p2bat = 0 And p3bat = 1 Then
                                            doliabatych = Frpor / 4
                                                If pprkrai = True Then
                                                    Xprsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xprbliggran)
                                                    Yprsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Yprbliggran)
                                                    X1b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Xprsrgran)
                                                    Y1b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Yprsrgran)
                                                    X3b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 5) + Xprsrgran)
                                                    Y3b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 5) + Yprsrgran)
                                                    OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                                                    OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                                                    OZzelkagdform.pXc2 = 0: OZzelkagdform.pYc2 = 0: OZzelkagdform.phc2 = 0: OZzelkagdform.pFrc2 = 0: OZzelkagdform.pGlc2 = 0
                                                    OZzelkagdform.Show
                                                    ElseIf pzentr = True Then
                                                            Xb = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xbliggran)
                                                            Yb = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ybliggran)
                                                            X1b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Xb)
                                                            Y1b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Yb)
                                                            X3b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Xb)
                                                            Y3b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Yb)
                                                            OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                                                            OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                                                            OZzelkagdform.pXc2 = 0: OZzelkagdform.pYc2 = 0: OZzelkagdform.phc2 = 0: OZzelkagdform.pFrc2 = 0: OZzelkagdform.pGlc2 = 0
                                                            OZzelkagdform.Show
                                                            ElseIf plevkrai = True Then
                                                                Xlevsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xlevbliggran)
                                                                Ylevsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ylevbliggran)
                                                                X3b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Xlevsrgran)
                                                                Y3b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Ylevsrgran)
                                                                X1b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Xlevsrgran)
                                                                Y1b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Ylevsrgran)
                                                                OZzelkagdform.pXc1 = X1b: OZzelkagdform.pYc1 = Y1b: OZzelkagdform.phc1 = hzc: OZzelkagdform.pFrc1 = Round(doliabatych * 2): OZzelkagdform.pGlc1 = Glpor
                                                                OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                                                                OZzelkagdform.pXc2 = 0: OZzelkagdform.pYc2 = 0: OZzelkagdform.phc2 = 0: OZzelkagdform.pFrc2 = 0: OZzelkagdform.pGlc2 = 0
                                                                OZzelkagdform.Show

                                                        Else
                                                End If
                ElseIf p1bat = 0 And p2bat = 1 And p3bat = 1 Then
                    doliabatych = Frpor / 4
                                                If pprkrai = True Then
                                                    Xprsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xprbliggran)
                                                    Yprsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Yprbliggran)
                                                    X2b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Xprsrgran)
                                                    Y2b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Yprsrgran)
                                                    X3b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Xprsrgran)
                                                    Y3b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Yprsrgran)
                                                    OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                                                    OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                                                    OZzelkagdform.pXc1 = 0: OZzelkagdform.pYc1 = 0: OZzelkagdform.phc1 = 0: OZzelkagdform.pFrc1 = 0: OZzelkagdform.pGlc1 = 0
                                                    OZzelkagdform.Show
                                                    ElseIf pzentr = True Then
                                                            Xb = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xbliggran)
                                                            Yb = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ybliggran)
                                                            X2b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Xb)
                                                            Y2b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Yb)
                                                            X3b = Round(Cos((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Xb)
                                                            Y3b = Round(Sin((Ygoltzc1 - 1500) / 100 * 6 * Pi / 180) * doliabatych + Yb)
                                                            OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                                                            OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                                                            OZzelkagdform.pXc1 = 0: OZzelkagdform.pYc1 = 0: OZzelkagdform.phc1 = 0: OZzelkagdform.pFrc1 = 0: OZzelkagdform.pGlc1 = 0
                                                            OZzelkagdform.Show
                                                            ElseIf plevkrai = True Then
                                                                Xlevsrgran = Round(Cos((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Xlevbliggran)
                                                                Ylevsrgran = Round(Sin((Ygoltzc + 0.001) / 100 * 6 * Pi / 180) * ((Glpor + 0.001) / 2) + Ylevbliggran)
                                                                X3b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Xlevsrgran)
                                                                Y3b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * doliabatych + Ylevsrgran)
                                                                X2b = Round(Cos((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Xlevsrgran)
                                                                Y2b = Round(Sin((Ygoltzc + 1500) / 100 * 6 * Pi / 180) * (doliabatych * 3) + Ylevsrgran)
                                                                OZzelkagdform.pXc3 = X3b: OZzelkagdform.pYc3 = Y3b: OZzelkagdform.phc3 = hzc: OZzelkagdform.pFrc3 = Round(doliabatych * 2): OZzelkagdform.pGlc3 = Glpor
                                                                OZzelkagdform.pXc2 = X2b: OZzelkagdform.pYc2 = Y2b: OZzelkagdform.phc2 = hzc: OZzelkagdform.pFrc2 = Round(doliabatych * 2): OZzelkagdform.pGlc2 = Glpor
                                                                OZzelkagdform.pXc1 = 0: OZzelkagdform.pYc1 = 0: OZzelkagdform.phc1 = 0: OZzelkagdform.pFrc1 = 0: OZzelkagdform.pGlc1 = 0
                                                                OZzelkagdform.Show

                                                        Else
                                                End If
    Else
End If
    Else
End If

End Sub
Function IZMERENIE(ByVal Xc As Single, ByVal Yc As Single, ByVal Xop As Single, ByVal Yop As Single, Dt, Ygolt) As Single
 Pi = 3.14159265358
 dx = Xc - Xop
 dy = Yc - Yop
 Dt1 = Int(Sqr(dx ^ 2 + dy ^ 2) + 0.001)
 A1 = Abs(Atn(dy / (dx + 0.001)) / Pi * 30) * 100
 If dx > 0 And dy > 0 Then Ygolt1 = CInt(A1)
 If dx < 0 And dy > 0 Then Ygolt1 = CInt(3000 - A1)
 If dx < 0 And dy < 0 Then Ygolt1 = CInt(3000 + A1)
 If dx > 0 And dy < 0 Then Ygolt1 = CInt(6000 - A1)
Dt = Dt1: Ygolt = Ygolt1
End Function

