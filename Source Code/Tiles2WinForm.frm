VERSION 5.00
Begin VB.Form Tiles2WinForm 
   BackColor       =   &H80000008&
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton XXit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton NNewGame 
      Caption         =   "&New Game"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label MoveL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "You did it in      moves!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "WELL DONE! YOU HAVE WON !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Tiles2WinForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Moves As Integer

Private Sub Form_Load()
  Moves = Tiles2Form.MoveCount
  MoveL = Moves
End Sub

Private Sub NNewGame_Click()
  Tiles2Form.MixUp
  Unload Me
End Sub

Private Sub XXit_Click()
End
End Sub
