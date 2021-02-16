VERSION 5.00
Begin VB.Form EditColourForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Colours"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2265
   Icon            =   "EditColourForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "EditColourForm.frx":030A
   ScaleHeight     =   3825
   ScaleWidth      =   2265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Message 
      Caption         =   "Message"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Background 
      Caption         =   "Background"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.VScrollBar BlueS 
      Height          =   1935
      Left            =   1800
      Max             =   255
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.VScrollBar GreenS 
      Height          =   1935
      Left            =   960
      Max             =   255
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.VScrollBar RedS 
      Height          =   1935
      Left            =   120
      Max             =   255
      TabIndex        =   1
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "G"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "R"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   135
   End
   Begin VB.Label InstructionL 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Choose background colour"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "EditColourForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CurrentOption As Byte

Private Sub Background_Click()
CurrentOption = 0
InstructionL.Caption = "Choose Background Colour"
End Sub

Private Sub BlueS_Change()
Select Case CurrentOption
Case 0: TileGameForm.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
Case 1: TileGameForm.MessageL.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
End Select
End Sub

Private Sub BlueS_Scroll()
Select Case CurrentOption
Case 0: TileGameForm.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
Case 1: TileGameForm.MessageL.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
End Select
End Sub

Private Sub greenS_Change()
Select Case CurrentOption
Case 0: TileGameForm.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
Case 1: TileGameForm.MessageL.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
End Select
End Sub

Private Sub greenS_Scroll()
Select Case CurrentOption
Case 0: TileGameForm.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
Case 1: TileGameForm.MessageL.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
End Select
End Sub

Private Sub Message_Click()
CurrentOption = 1
InstructionL.Caption = "Choose Message Colour"
End Sub

Private Sub OK_Click()
   Unload Me
End Sub

Private Sub redS_Change()
Select Case CurrentOption
Case 0: TileGameForm.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
Case 1: TileGameForm.MessageL.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
End Select
End Sub

Private Sub redS_Scroll()
Select Case CurrentOption
Case 0: TileGameForm.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
Case 2: TileGameForm.MessageL.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
End Select
End Sub
