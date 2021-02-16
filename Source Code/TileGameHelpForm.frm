VERSION 5.00
Begin VB.Form TileGameHelpForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to play 'Tiles'"
   ClientHeight    =   3075
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton OK2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   120
      Picture         =   "TileGameHelpForm.frx":0000
      ScaleHeight     =   2415
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label HelpL 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"TileGameHelpForm.frx":030A
         Height          =   2355
         Left            =   660
         TabIndex        =   1
         Top             =   120
         Width           =   2775
      End
   End
End
Attribute VB_Name = "TileGameHelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OK2_Click()
Unload Me
End Sub
