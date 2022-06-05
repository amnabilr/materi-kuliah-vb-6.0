VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   2400
      TabIndex        =   6
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   2400
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MESSAGE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "INPUT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   2
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "INPUT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INPUT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = InputBox("Type Your First Name", "input", vbOKOnly, 600, 500)
Text1.Text = a
End Sub

Private Sub Command2_Click()
b = InputBox("Type Your Middle Name", "input", vbOKOnly, 600, 500)
Text2.Text = b
End Sub

Private Sub Command3_Click()
c = InputBox("Type Your Last Name", "input", vbOKOnly, 600, 500)
Text3.Text = c
End Sub

Private Sub Command4_Click()
d = MsgBox("Welcome ! " & Text1.Text & Text2.Text & Text3.Text)
e = MsgBox("Do you want to quit ?", vbYesNo + vbQuestion)
End Sub

Private Sub Text1_Change()
Caption = Command1
End Sub

Private Sub Text2_Change()
Caption = Command2
End Sub

Private Sub Text3_Change()
Caption = Command3
End Sub
