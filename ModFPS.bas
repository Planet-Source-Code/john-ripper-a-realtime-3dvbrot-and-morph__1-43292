Attribute VB_Name = "ModFPS"
Option Explicit

Dim Tick            As Long
Dim ElapsedTicks    As Long
Dim LastTick        As Long
Dim TickCounter     As Long
Dim FrameCounter    As Long

Private FPS          As Long

Private Sub CalcTick()
    Tick = timeGetTime
    ElapsedTicks = Tick - LastTick
    LastTick = Tick
End Sub
Private Sub CalcFPS()
    TickCounter = TickCounter + ElapsedTicks
    
    If TickCounter > 1000 Then
        FPS = 1000 * FrameCounter \ TickCounter
        FrameCounter = 0
        TickCounter = 0
    End If
    
    FrameCounter = FrameCounter + 1
End Sub

Public Sub ShowFPS()
Dim lRet As Boolean
    CalcTick
    CalcFPS
    If isGraphFPS = True Then
        CleanHorizontalAsc
        SDC = frmSprites.PicSPRascii.hdc
        DDC = frmBack.PicASCHorizontalWork.hdc
        WriteMyAscii SDC, DDC, rWidthASC, Str(FPS), 0, 0, False, 2
        BlitHorizontalAsc
    Else
        frmMain.lblFPS.Caption = FPS
    End If
End Sub


