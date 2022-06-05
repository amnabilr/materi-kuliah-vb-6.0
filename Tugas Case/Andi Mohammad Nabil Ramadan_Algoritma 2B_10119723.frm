VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nama Planet di Tata Surya"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim no_planet As Integer
no_planet = Val(Text1.Text)
Select Case no_planet
Case 1
MsgBox "Planet Merkurius"
Case 2
MsgBox "Planet Venus"
Case 3
MsgBox "Planet Bumi"
Case 4
MsgBox "Planet Mars"
Case 5
MsgBox "Planet Jupiter"
Case 6
MsgBox "Planet Saturnus"
Case 7
MsgBox "Planet Neptunus"
Case 8
MsgBox "Planet Uranus"
Case 9
MsgBox "Planet Pluto"
End Select
End Sub

