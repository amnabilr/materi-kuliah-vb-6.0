VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12285
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "UANG MUKA 20 % DR HRG RUMAH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "HRG RUMAH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TYPE RUMAH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Dim HARGA As Single
HARGA = Combo1.ItemData(Combo1.ListIndex)
Text1.Text = Format(HARGA, "Rp. #.###.###.###,##")
Text2.Text = Format(0.2 * HARGA, "Rp. #.###.###.###,##")
Text2.SetFocus
Form1.BackColor = vbMagenta
Text1.BackColor = vbYellow
Text2.BackColor = vbYellow
End Sub
Private Sub Command1_Click()
End
End Sub
Private Sub Form_Load()
Combo1.List(0) = "Type 21"
Combo1.List(1) = "Type 36"
Combo1.List(2) = "Type 40"
Combo1.List(3) = "Type 72"
Combo1.List(4) = "Type 108"
Combo1.ItemData(0) = 45000000
Combo1.ItemData(1) = 75000000
Combo1.ItemData(2) = 95000000
Combo1.ItemData(3) = 105000000
Combo1.ItemData(4) = 165000000
End Sub

