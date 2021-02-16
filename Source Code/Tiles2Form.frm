VERSION 5.00
Begin VB.Form Tiles2Form 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tiles"
   ClientHeight    =   2895
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   3240
   Icon            =   "Tiles2Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Tiles2Form.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   3
   ScaleMode       =   0  'User
   ScaleWidth      =   3
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Tile 
      BackColor       =   &H0000FFFF&
      Caption         =   "5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   965
      Index           =   5
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Move 5"
      Top             =   960
      Width           =   1080
   End
   Begin VB.CommandButton Tile 
      BackColor       =   &H0000FFFF&
      Caption         =   "8"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   965
      Index           =   8
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Move 8 "
      Top             =   1930
      Width           =   1080
   End
   Begin VB.CommandButton Tile 
      BackColor       =   &H0000FFFF&
      Caption         =   "7"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   965
      Index           =   7
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Move 7"
      Top             =   1930
      Width           =   1080
   End
   Begin VB.CommandButton Tile 
      BackColor       =   &H0000FFFF&
      Caption         =   "6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   965
      Index           =   6
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Move 6"
      Top             =   965
      Width           =   1080
   End
   Begin VB.CommandButton Tile 
      BackColor       =   &H0000FFFF&
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   965
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move 4"
      Top             =   965
      Width           =   1080
   End
   Begin VB.CommandButton Tile 
      BackColor       =   &H0000FFFF&
      Caption         =   "3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   965
      Index           =   3
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Move 3"
      Top             =   0
      Width           =   1080
   End
   Begin VB.CommandButton Tile 
      BackColor       =   &H0000FFFF&
      Caption         =   "2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   965
      Index           =   2
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Move 2"
      Top             =   0
      Width           =   1080
   End
   Begin VB.CommandButton Tile 
      BackColor       =   &H0000FFFF&
      Caption         =   "1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   965
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Move 1"
      Top             =   0
      Width           =   1080
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu NewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu Colours 
         Caption         =   "&Colours..."
      End
      Begin VB.Menu Xit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu HowToPlay 
         Caption         =   "How to &Play"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Tiles2Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SpacePos As Byte 'stores position of space
Public SpacePos2 As Byte 'stores previous position of space
Dim TilePos(1 To 8) As Byte 'stores positions of tiles
Public n As Integer 'used as counter
Public c As Integer 'used as counter
Public Index2 As Byte 'stores indexes of tiles
Public time As Byte 'used to choose tile speed
Public MoveCount As Integer 'counts the number of moves taken

Private Sub About_Click()
  AboutTiles2Form.Visible = True
End Sub

Private Sub Colours_Click()
Tiles2ColoursForm.Visible = True
End Sub

Private Sub Form_Load()
Show
Randomize 'start random number generator
For n = 1 To 8
TilePos(n) = n 'sets positions of tiles
Next
SpacePos = 9 'space is in position 9
MixUp 'MixUpTheTiles
End Sub

Sub MixUp()
MoveCount = 0
time = 100
MixingUpForm.Visible = True
For n = 1 To 8
  Tiles2Form.Tile(n).Enabled = False
Next
For c = 1 To 1000
  Index2 = Int(Rnd * 8) + 1
  Tile(Index2).Enabled = True
  Tile(Index2).Value = True
  Tile(Index2).Enabled = False
  MixingUpForm.ProgressBar1.Value = c
Next
For n = 1 To 8
  Tiles2Form.Tile(n).Enabled = True
Next
MixingUpForm.Visible = False
time = 1
'sndPlaySound "A:\SoundFX\Tiles\MoveIt.wav", &H1
End Sub

Private Sub HowToPlay_Click()
   Tile2HelpForm.Visible = True
End Sub

Private Sub NewGame_Click()
MixUp
End Sub

Private Sub Tile_Click(Index As Integer)
Index2 = Index
Select Case SpacePos
       Case 1: Select Case TilePos(Index)
                      Case 2: MoveLeft
                      Case 4: MoveUp
                      Case Else: Exit Sub
               End Select
       Case 2: Select Case TilePos(Index)
                      Case 1: MoveRight
                      Case 3: MoveLeft
                      Case 5: MoveUp
                      Case Else: Exit Sub
               End Select
       Case 3: Select Case TilePos(Index)
                      Case 2: MoveRight
                      Case 6: MoveUp
                      Case Else: Exit Sub
               End Select
       Case 4: Select Case TilePos(Index)
                      Case 1: MoveDown
                      Case 7: MoveUp
                      Case 5: MoveLeft
                      Case Else: Exit Sub
               End Select
       Case 5: Select Case TilePos(Index)
                      Case 2: MoveDown
                      Case 8: MoveUp
                      Case 4: MoveRight
                      Case 6: MoveLeft
                      Case Else: Exit Sub
               End Select
       Case 6: Select Case TilePos(Index)
                      Case 3: MoveDown
                      Case 9: MoveUp
                      Case 5: MoveRight
                      Case Else: Exit Sub
               End Select
       Case 7: Select Case TilePos(Index)
                      Case 4: MoveDown
                      Case 8: MoveLeft
                      Case Else: Exit Sub
               End Select
       Case 8: Select Case TilePos(Index)
                      Case 9: MoveLeft
                      Case 7: MoveRight
                      Case 5: MoveDown
                      Case Else: Exit Sub
               End Select
       Case 9: Select Case TilePos(Index)
                      Case 8: MoveRight
                      Case 6: MoveDown
                      Case Else: Exit Sub
               End Select
End Select
SpacePos2 = SpacePos 'save last spacepos
SpacePos = TilePos(Index2) 'set new pos of space as old pos of tile
TilePos(Index2) = SpacePos2 'set new pos of tile as old pos of space
If MixingUpForm.Visible = True Then Exit Sub
MoveCount = MoveCount + 1
For n = 1 To 8
Select Case TilePos(n)
       Case n:
       Case Else: Exit Sub
End Select
Next
Tiles2WinForm.Visible = True
End Sub

Sub MoveUp()
'If MixingUpForm.Visible = False Then 'sndPlaySound "A:\SoundFX\Tiles\Slide1.wav", &H1
For n = 1 To 1000 / time
Tile(Index2).Top = Tile(Index2).Top - 0.001 * time 'move up
Next
Tile(Index2).Top = Int(Tile(Index2).Top + 0.1) 'line up nicely
End Sub

Sub MoveDown()
'If MixingUpForm.Visible = False Then sndPlaySound "A:\SoundFX\Tiles\Slide2.wav", &H1
For n = 1 To 1000 / time
Tile(Index2).Top = Tile(Index2).Top + 0.001 * time 'move down
Next
Tile(Index2).Top = Int(Tile(Index2).Top + 0.1) 'line up nicely
End Sub

Sub MoveLeft()
'If MixingUpForm.Visible = False Then sndPlaySound "A:\SoundFX\Tiles\Slide3.wav", &H1
For n = 1 To 1000 / time
Tile(Index2).Left = Tile(Index2).Left - 0.001 * time 'move left
Next
Tile(Index2).Left = Int(Tile(Index2).Left + 0.1) 'line up nicely
End Sub

Sub MoveRight()
'If MixingUpForm.Visible = False Then sndPlaySound "A:\SoundFX\Tiles\Slide4.wav", &H1
For n = 1 To 1000 / time
Tile(Index2).Left = Tile(Index2).Left + 0.001 * time 'move right
Next
Tile(Index2).Left = Int(Tile(Index2).Left + 0.1) 'line up nicely
End Sub

Private Sub Xit_Click()
End
End Sub
