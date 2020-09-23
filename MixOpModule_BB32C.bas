Attribute VB_Name = "MixOpModule_BB32C"
Public Const blVertical As Boolean = True, blHorizontal As Boolean = False
Public FStep(21) As Long, StepC As Byte
Public R(21) As Long, G(21) As Long, B(21) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ObjFade(OB As Object, ByVal Direction As Boolean, Optional BackColorR As Byte, Optional BackColorG As Byte, Optional BackColorB As Byte)  'RGB tones and FadeSteps are located in their respective variables
Dim RStep As Single, GStep As Single, BStep As Single, J As Integer, CurR As Single, CurG As Single, CurB As Single
numGetFadingCount
If StepC = 0 Then
 MsgBox "No fade points specified!", vbExclamation + vbApplicationModal
 Exit Sub
End If
 If Direction = blVertical Then FStep(StepC) = OB.ScaleHeight Else FStep(StepC) = OB.ScaleWidth
 R(StepC) = BackColorR
 G(StepC) = BackColorG
 B(StepC) = BackColorB
 R(0) = BackColorR
 G(0) = BackColorG
 B(0) = BackColorB
 CurR = R(0)
 CurG = G(0)
 CurB = B(0)
 For I = 0 To StepC
 Select Case Direction
 Case blVertical
  If FStep(I + 1) - FStep(I) = 0 Then GoTo NotOverflow
  RStep = (R(I + 1) - R(I)) / (FStep(I + 1) - FStep(I))
  GStep = (G(I + 1) - G(I)) / (FStep(I + 1) - FStep(I))
  BStep = (B(I + 1) - B(I)) / (FStep(I + 1) - FStep(I))
  For J = FStep(I) To FStep(I + 1)
   OB.Line (0, J)-(OB.ScaleWidth, J), RGB(Fix(CurR), Fix(CurG), Fix(CurB))
   CurR = CurR + RStep
   CurG = CurG + GStep
   CurB = CurB + BStep
  Next J
NotOverflow:
  CurR = R(I + 1)
  CurG = G(I + 1)
  CurB = B(I + 1)
 Case blHorizontal
  If FStep(I + 1) - FStep(I) = 0 Then GoTo NotOverflowH
  RStep = (R(I + 1) - R(I)) / (FStep(I + 1) - FStep(I))
  GStep = (G(I + 1) - G(I)) / (FStep(I + 1) - FStep(I))
  BStep = (B(I + 1) - B(I)) / (FStep(I + 1) - FStep(I))
  For J = FStep(I) To FStep(I + 1)
   OB.Line (J, 0)-(J, OB.ScaleHeight), RGB(Fix(CurR), Fix(CurG), Fix(CurB))
   CurR = CurR + RStep
   CurG = CurG + GStep
   CurB = CurB + BStep
  Next J
NotOverflowH:
  CurR = R(I + 1)
  CurG = G(I + 1)
  CurB = B(I + 1)
End Select
 Next I
 On Error GoTo 0
 R(0) = 0
 G(0) = 0
 B(0) = 0
 R(StepC) = 0
 G(StepC) = 0
 B(StepC) = 0
 FStep(StepC) = 0
 StepC = 0
End Sub

Public Sub numGetFadingCount()
For I = 20 To 1 Step -1
 If FStep(I - 1) <> 0 Then
  StepC = I
  Exit For
 End If
Next I
End Sub

Function LongToR(Color As Long) As Integer
Dim tg&, tc&
tc = Color
tg = tc \ 65536
tc = tc - (65536 * tg)
tg = tc \ 256
LongToR = tc - (256 * tg)
End Function

Function LongToG(Color As Long) As Integer
Dim tg&, tc&
tc = Color
tg = tc \ 65536
tc = tc - (65536 * tg)
LongToG = tc \ 256
End Function

Function LongToB(Color As Long) As Integer
LongToB = Color \ 65536
End Function
