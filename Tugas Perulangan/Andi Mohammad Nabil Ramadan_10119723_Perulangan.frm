VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "For Next"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   7
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Do While Loop"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do Until"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   3615
   End
   Begin VB.ListBox List3 
      Height          =   2595
      Left            =   9600
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   4920
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = 1
Do Until x > Val(Text1.Text)
List1.AddItem (x)
x = x + 1
Loop
End Sub

Private Sub Command2_Click()
x = 1
Do While x <= Val(Text1.Text)
List2.AddItem (x)
x = x + 1
Loop
End Sub

Private Sub Command3_Click()
For x = 1 To Val(Text1.Text)
angka = ""
For b = 1 To x
angka = angka & b
Next
List3.AddItem angka
Next x
End Sub

