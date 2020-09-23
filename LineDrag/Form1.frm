VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Polygonal Select"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Width           =   6060
      Begin VB.Line ALine 
         DrawMode        =   6  'Mask Pen Not
         Index           =   0
         X1              =   0
         X2              =   128
         Y1              =   -0.667
         Y2              =   -0.667
      End
   End
   Begin VB.Menu mnuLoad 
      Caption         =   "&Load Region"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save Region"
   End
   Begin VB.Menu mnuClear 
      Caption         =   "&Clear Regions"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LMode As Boolean
Dim Px() As Single, Py() As Single
Dim CoorCount As Integer

Private Sub Form_Resize()
Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuClear_Click()
' Clears the coordinates from memory

For a = 1 To ALine.UBound
Unload ALine(a)
Next a

ReDim Px(0)
ReDim Py(0)
CoorCount = 0

End Sub

Private Sub mnuLoad_Click()
Dim Dx As Single, Dy As Single
' Load the coordinates

' Check if the file is existing
If Dir$(App.Path + "\Region.txt") = "" Then
  MsgBox "No Region file found."
  Exit Sub
End If

' Open it and read it
Open App.Path + "\Region.txt" For Input As #1
  Do Until EOF(1)
    Input #1, Dx, Dy
    AddPoint Dx, Dy
  Loop
Close #1


' Apply the loaded coordinates
For a = 1 To CoorCount
  If a < CoorCount Then
    AddLine Px(a), Py(a)
    ALine(ALine.UBound).X2 = Px(a + 1)
    ALine(ALine.UBound).Y2 = Py(a + 1)
  End If
Next a


End Sub

Private Sub mnuSave_Click()
' Save the coordinates of the region in a file.

Open App.Path + "\Region.txt" For Output As #1
  For a = 1 To CoorCount
    Write #1, Px(a), Py(a)
  Next a
Close #1

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If LMode = True Then
  ALine(ALine.UBound).X2 = x
  ALine(ALine.UBound).Y2 = y
End If
End Sub

Sub AddLine(wX As Single, wY As Single)
' Adds a line to the polygon.
Load ALine(ALine.Count)
ALine(ALine.UBound).Visible = True
ALine(ALine.UBound).X1 = wX
ALine(ALine.UBound).Y1 = wY
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If LMode = False And Button = 1 Then
  LMode = True
  AddLine x, y
  AddPoint x, y
ElseIf LMode = True And Button = 1 Then
  AddLine x, y
  AddPoint x, y
ElseIf LMode = True And Button = 2 Then
  LMode = False
  Unload ALine(ALine.UBound)
End If
End Sub

Sub AddPoint(x, y)

' Add another coordinate to the count
CoorCount = CoorCount + 1

' Resize the array and add another coordinate
ReDim Preserve Px(CoorCount)
Px(UBound(Px)) = x

' Do the same thing for the Y coordinate
ReDim Preserve Py(CoorCount)
Py(UBound(Py)) = y

End Sub
