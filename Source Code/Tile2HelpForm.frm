VERSION 5.00
Begin VB.Form Tile2HelpForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to play tiles"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2025
   ControlBox      =   0   'False
   Icon            =   "Tile2HelpForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   2025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Done 
      Caption         =   "&Done"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Tile2HelpForm.frx":000C
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Tile2HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Done_Click()
  Unload Me
End Sub
