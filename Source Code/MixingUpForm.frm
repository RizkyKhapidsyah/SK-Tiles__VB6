VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MixingUpForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mixing Up..."
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   630
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   1000
   End
End
Attribute VB_Name = "MixingUpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

