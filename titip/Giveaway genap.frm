VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SELESAI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "HASIL INPUTAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "INPUT DATA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0080C0FF&
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "CLICK FORM LIHAT NAMA DAN NILAI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "NILAI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NAMA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N(10) As String * 15
Dim a(10), i As Integer
Private Sub Command1_Click()
For i = 0 To 2
N(i) = InputBox("Masukan data Nama ", "DATA NAMA AKAN DISIMPAN DLM ARRAY")
a(i) = InputBox("Masukan data Nilai ", "DATA NILAI JG DISIMPAN DLM ARRAY")
If a(i) = vbCancel Then
Exit For
End If
Next i
End Sub
Private Sub Command2_Click()
For i = 0 To 2
Label1.Caption = Label1.Caption & " " & N(i)
Label2.Caption = Label2.Caption & "        " & a(i)
Next i
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Form_Click()
For i = 0 To 2
List1.List(i) = N(i) & "  " & a(i)
Next i
End Sub
