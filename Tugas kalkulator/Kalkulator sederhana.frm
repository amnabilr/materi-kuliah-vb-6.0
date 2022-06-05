VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "START"
      Height          =   495
      Left            =   10680
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   10680
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "INPUT"
      Height          =   495
      Left            =   10680
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   10
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "HASIL"
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "BILANGAN KEDUA"
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "BILANGAN PERTAMA"
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "KALKULATOR SEDERHANA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3.Text = Val(Text1) + Val(Text2)
End Sub

Private Sub Command2_Click()
Text3.Text = Val(Text1) - Val(Text2)
End Sub

Private Sub Command3_Click()
Text3.Text = Val(Text1) * Val(Text2)
End Sub

Private Sub Command4_Click()
Text3.Text = Val(Text1) / Val(Text2)
End Sub

Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text1.SetFocus

End Sub

Private Sub Form_Load()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End Sub

