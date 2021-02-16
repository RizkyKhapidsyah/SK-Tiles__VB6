VERSION 5.00
Begin VB.Form TileGameForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Tiles - by Simon Price"
   ClientHeight    =   7245
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   11880
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PB 
      BackColor       =   &H80000007&
      Height          =   7095
      Left            =   360
      ScaleHeight     =   3
      ScaleMode       =   0  'User
      ScaleWidth      =   3
      TabIndex        =   0
      Top             =   2
      Width           =   10935
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   9
         Left            =   7250
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   9
         Top             =   4667
         Width           =   3625
      End
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   8
         Left            =   3625
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   8
         Top             =   4667
         Width           =   3625
      End
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   7
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   7
         Top             =   4667
         Width           =   3625
      End
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   5
         Left            =   3625
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   6
         Top             =   2345
         Width           =   3625
      End
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   3
         Left            =   7250
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   5
         Top             =   0
         Width           =   3625
      End
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   6
         Left            =   7250
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   4
         Top             =   2345
         Width           =   3625
      End
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   4
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   3
         Top             =   2345
         Width           =   3625
      End
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   2
         Left            =   3625
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   2
         Top             =   0
         Width           =   3625
      End
      Begin VB.PictureBox Chunk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2345
         Index           =   1
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   1
         TabIndex        =   1
         Top             =   0
         Width           =   3625
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu ChooseAPicture 
         Caption         =   "&Choose a &Picture"
      End
      Begin VB.Menu Xit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "TileGameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SpacePos As Byte
'Public Index As Byte
Public Index2 As Byte
Public TileIsMoving As Boolean
Public StopNow As Byte
Public n As Byte

Private Sub ChooseAPicture_Click() 'Goes to other form
PictureViewerForm.Visible = True
End Sub

Private Sub Chunk_Click(Index As Integer) 'when any chunk is clicked
Print "Chunk " & Index & " clicked"
Index2 = Index
If TileIsMoving Then
TileGameForm.Chunk(1).PaintPicture PictureViewerForm.PB.Picture, 0, 0 're-paints chunks
TileGameForm.Chunk(2).PaintPicture PictureViewerForm.PB.Picture, -1, 0
TileGameForm.Chunk(3).PaintPicture PictureViewerForm.PB.Picture, -2, 0
TileGameForm.Chunk(4).PaintPicture PictureViewerForm.PB.Picture, 0, -1
TileGameForm.Chunk(5).PaintPicture PictureViewerForm.PB.Picture, -1, -1
TileGameForm.Chunk(6).PaintPicture PictureViewerForm.PB.Picture, -2, -1
TileGameForm.Chunk(7).PaintPicture PictureViewerForm.PB.Picture, 0, -1.99
TileGameForm.Chunk(8).PaintPicture PictureViewerForm.PB.Picture, -1, -1.99
TileGameForm.Chunk(9).PaintPicture PictureViewerForm.PB.Picture, -2, -1.99
Exit Sub 'if busy, ignore user
End If
TileIsMoving = True
Select Case SpacePos 'where is the space?
Case 1: Select Case Index 'what chunk was clicked?
                      Case 2: MoveLeft 'move is possible, so move
                      Case 4: MoveUp
                      Case Else: Exit Sub 'move not possible, so do not move
               End Select
Case 2: Select Case Index
                      Case 1: MoveRight
                      Case 5: MoveUp
                      Case 3: MoveLeft
                      Case Else: Exit Sub
               End Select
Case 3: Select Case Index
                      Case 2: MoveRight
                      Case 6: MoveUp
                      Case Else: Exit Sub
               End Select
Case 4: Select Case Index
                      Case 2: MoveDown
                      Case 7: MoveUp
                      Case 5: MoveLeft
                      Case Else: Exit Sub
               End Select
Case 5: Select Case Index
                      Case 2: MoveDown
                      Case 4: MoveRight
                      Case 6: MoveLeft
                      Case 8: MoveUp
                      Case Else: Exit Sub
               End Select
Case 6: Select Case Index
                      Case 3: MoveDown
                      Case 9: MoveUp
                      Case 5: MoveRight
                      Case Else: Exit Sub
               End Select
Case 7: Select Case Index
                      Case 8: MoveLeft
                      Case 4: MoveDown
                      Case Else: Exit Sub
               End Select
Case 8: Select Case Index
                      Case 7: MoveRight
                      Case 9: MoveLeft
                      Case 5: MoveDown
                      Case Else: Exit Sub
               End Select
Case 9: Select Case Index
                      Case 6: MoveDown
                      Case 8: MoveRight
                      Case Else: Exit Sub
               End Select
End Select
     'This bit...
Index = SpacePos    '...switches index...
SpacePos = Index2
Chunk(SpacePos).Visible = False '...and...
Chunk(Index2).Index = Index '...spacepos
Chunk(Index).Visible = True
TileGameForm.Chunk(1).PaintPicture PictureViewerForm.PB.Picture, 0, 0 're-paints chunks
TileGameForm.Chunk(2).PaintPicture PictureViewerForm.PB.Picture, -1, 0
TileGameForm.Chunk(3).PaintPicture PictureViewerForm.PB.Picture, -2, 0
TileGameForm.Chunk(4).PaintPicture PictureViewerForm.PB.Picture, 0, -1
TileGameForm.Chunk(5).PaintPicture PictureViewerForm.PB.Picture, -1, -1
TileGameForm.Chunk(6).PaintPicture PictureViewerForm.PB.Picture, -2, -1
TileGameForm.Chunk(7).PaintPicture PictureViewerForm.PB.Picture, 0, -1.99
TileGameForm.Chunk(8).PaintPicture PictureViewerForm.PB.Picture, -1, -1.99
TileGameForm.Chunk(9).PaintPicture PictureViewerForm.PB.Picture, -2, -1.99
TileIsMoving = False
End Sub

Private Sub Form_Load()
SpacePos = 9
Chunk(SpacePos).Visible = False
End Sub

Sub MoveUp()
Select Case Index2
       Case 7: StopNow = 99
       Case 8: StopNow = 99
       Case 9: StopNow = 99
       Case Else: StopNow = 100
End Select
For n = 1 To StopNow
Chunk(Index2).Top = Chunk(Index2).Top - 0.01
Next
End Sub

Sub MoveDown()
Select Case Index2
       Case 4: StopNow = 99
       Case 5: StopNow = 99
       Case 6: StopNow = 99
       Case Else: StopNow = 100
End Select
For n = 1 To StopNow
Chunk(Index2).Top = Chunk(Index2).Top + 0.01
Next
End Sub

Sub MoveLeft()
For n = 1 To 100
Chunk(Index2).Left = Chunk(Index2).Left - 0.01
Next
End Sub

Sub MoveRight()
For n = 1 To 100
Chunk(Index2).Left = Chunk(Index2).Left + 0.01
Next
End Sub
