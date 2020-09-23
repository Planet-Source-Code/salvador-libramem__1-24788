Attribute VB_Name = "DrawPie"
Option Explicit
   Private Const PI As Double = 3.14159265359
   Private Const CircleEnd As Double = -2 * PI
   Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
   Dim coord As POINT_TYPE
   Dim retval As Long, sinact, mix, miy
   Public mins As Integer
   Type POINT_TYPE
        X As Long
        Y As Long
   End Type

Public Sub DrawPiePiece(myform As Form, lColor As Long, fStart As Double, fEnd As Double)
    Dim dStart As Double
    Dim dEnd As Double
    myform.FillColor = lColor
    myform.FillStyle = 0
    dStart = fStart * (CircleEnd / 100)
    dEnd = fEnd * (CircleEnd / 100)
    myform.Circle (20, 20), 20, , dStart, dEnd
End Sub

Function inact()
     retval = GetCursorPos(coord)
     If mix <> coord.X Then mix = coord.X: sinact = 0: Exit Function
     If miy <> coord.Y Then miy = coord.Y: sinact = 0: Exit Function
     sinact = sinact + 1
     If sinact > 60 Then
        sinact = 0
        mins = mins + 1
      End If
End Function
