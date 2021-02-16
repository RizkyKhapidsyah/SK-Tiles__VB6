VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Tiles2ColoursForm 
   Caption         =   "Select Tile Colour"
   ClientHeight    =   3432
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   5328
   Icon            =   "Tiles2ColoursForm.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   3432
   ScaleWidth      =   5328
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OK3 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      ToolTipText     =   "Accept this colour scheme"
      Top             =   3000
      Width           =   1575
   End
   Begin ComctlLib.Slider RedS 
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Choose amount of red"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4255
      _ExtentY        =   720
      _Version        =   327680
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "Tiles2ColoursForm.frx":030A
      LargeChange     =   25
      Max             =   255
      SelectRange     =   -1  'True
      SelStart        =   255
      TickFrequency   =   51
      Value           =   255
   End
   Begin VB.CommandButton Preview 
      BackColor       =   &H0000FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      MouseIcon       =   "Tiles2ColoursForm.frx":0624
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Preview of tile "
      Top             =   240
      Width           =   855
   End
   Begin VB.OptionButton Rainbow 
      Caption         =   "Rainbow"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.OptionButton CustomO 
      Caption         =   "Create a Tile Colour"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton SystemO 
      Caption         =   "Use System Colours"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton DefaultO 
      Caption         =   "Default colour (yellow)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1935
   End
   Begin ComctlLib.Slider GreenS 
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Choose amount of green"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4255
      _ExtentY        =   720
      _Version        =   327680
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "Tiles2ColoursForm.frx":0776
      LargeChange     =   25
      Max             =   255
      SelectRange     =   -1  'True
      SelStart        =   255
      TickFrequency   =   51
      Value           =   255
   End
   Begin ComctlLib.Slider BlueS 
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "Choose amount of blue"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4255
      _ExtentY        =   720
      _Version        =   327680
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "Tiles2ColoursForm.frx":0A90
      LargeChange     =   25
      Max             =   255
      SelectRange     =   -1  'True
      SelStart        =   255
      TickFrequency   =   51
      Value           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Blue"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Green"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Red"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   2280
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Tiles2ColoursForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public n As Byte
Public nn As Byte

Private Sub CustomO_Click()
RedS.Visible = True

GreenS.Visible = True
BlueS.Visible = True
End Sub

Sub UpdateColour()
Preview.BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
For n = 1 To 8
   Tiles2Form.Tile(n).BackColor = RGB(RedS.Value, GreenS.Value, BlueS.Value)
Next
End Sub

Private Sub DefaultO_Click()
Preview.Visible = True
RedS.Visible = False
GreenS.Visible = False
BlueS.Visible = False
Preview.BackColor = &HFFFF&
For n = 1 To 8
   Tiles2Form.Tile(n).BackColor = &HFFFF&
Next
End Sub

Private Sub OK3_Click()
'sndPlaySound "A:\SoundFX\Tiles\GoodChoice.wav", 0
Unload Me
End Sub

Private Sub Rainbow_Click()
Preview.Visible = True
Preview.Visible = False
Tiles2Form.Tile(1).BackColor = QBColor(1)
Tiles2Form.Tile(2).BackColor = QBColor(2)
Tiles2Form.Tile(3).BackColor = QBColor(3)
Tiles2Form.Tile(4).BackColor = QBColor(4)
Tiles2Form.Tile(5).BackColor = QBColor(5)
Tiles2Form.Tile(6).BackColor = QBColor(9)
Tiles2Form.Tile(7).BackColor = QBColor(10)
Tiles2Form.Tile(8).BackColor = QBColor(14)
End Sub

Private Sub RedS_Change()
UpdateColour
End Sub

Private Sub GreenS_Change()
UpdateColour
End Sub

Private Sub BlueS_Change()
UpdateColour
End Sub

Private Sub SystemO_Click()
Preview.Visible = True
RedS.Visible = False
GreenS.Visible = False
BlueS.Visible = False
Preview.BackColor = &H8000000F
For n = 1 To 8
   Tiles2Form.Tile(n).BackColor = &H8000000F
Next
End Sub
