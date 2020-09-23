Attribute VB_Name = "ModScenes"
Option Explicit
Dim lRet As Boolean


Public isGraphFPS               As Boolean
Public isBackGround             As Boolean
Public isSleep                  As Boolean
Public isPixelDOTT              As Boolean
Public isQZsortDOTT             As Boolean
Public isPixelSingleColorDOTT   As Boolean
Public isRedPixelDOTT           As Boolean
Public isGreenPixelDOTT         As Boolean
Public isBluePixelDOTT          As Boolean
Public FogValueDOTT             As Integer
Public isAlphaEffect            As Boolean
Public isSolidSprite            As Boolean
Public isShadowedSprites        As Boolean
Public isBenny                  As Boolean
Public isQZsortWIRE             As Boolean
Public isHidenFacesWIRE         As Boolean
Public isNoDiagonals            As Boolean
Public isWireSingleColor        As Boolean
Public isRedPixelWIRE           As Boolean
Public isGreenPixelWIRE         As Boolean
Public isBluePixelWIRE          As Boolean
Public FogValueWIRE             As Integer

Public Sub FirstScene()
    ReadFileMesh App.Path & "\Meshes\Bola314_1.dVB", Points, Faces
    
    Do Until lRet = True
        
        If isAlphaEffect = True And Not isPixelDOTT And Not isSolidSprite Then
            DoSceneA Points, Faces, isBackGround, isQZsortDOTT, Not isPixelDOTT, frmSprites.PicSPR3d.hdc, False, Not isPixelSingleColorDOTT, , SpeedXangle, SpeedYangle, SpeedZangle, frmSprites.PicPlanets.hdc, True, isAlphaEffect
        Else
            DoSceneA Points, Faces, isBackGround, isQZsortDOTT, Not isPixelDOTT, frmSprites.PicSPR3d.hdc, False, Not isPixelSingleColorDOTT, , SpeedXangle, SpeedYangle, SpeedZangle, , , isAlphaEffect
        End If
        Blit3D
        If isSleep = True Then
            Sleep 16
        End If
        ShowFPS
        DoEvents
    Loop
End Sub

Public Sub RotateSolarSystem()
Dim i As Integer

    ReadFileMesh App.Path & "\Meshes\Planets.dVB", Points, Faces
    For i = 0 To 8
        Points(i).Aux = i
    Next i
    Xangle = 0
    Yangle = 0
    Zangle = 0
    
    lRet = False
    Do Until lRet = True
        If isBackGround = True Then
            SDC = frmSprites.PicMainBackGround.hdc
        Else
            SDC = frmBack.PicMainClean.hdc
        End If
        DDC = frmBack.PicMainWork.hdc
        BitBlt DDC, 0, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY
        DoRotation Points, isQZsortDOTT, True, frmSprites.PicPlanets.hdc, False, , , SpeedXangle, SpeedYangle, SpeedZangle, frmSprites.PicPlanets.hdc
        
        Blit3D
        If isSleep = True Then
            Sleep 16
        End If
        ShowFPS
        DoEvents
    Loop
End Sub

Public Sub ApplyEngineWire(MeshFile As String)
Dim i As Integer
Dim hDCSPR As Long
    PrepareMorphing Points, FinishPoints, MeshFile, True
    
    For i = 0 To frmMain.ImgPIX.Count - 1
        frmMain.ImgPIX(i).Enabled = False
    Next i
    
    For i = 0 To frmMain.ImgWIRE.Count - 1
        frmMain.ImgWIRE(i).Enabled = False
    Next i
    
    lRet = False
    Do Until lRet = True
        If isBackGround = True Then
            SDC = frmSprites.PicMainBackGround.hdc
        Else
            SDC = frmBack.PicMainClean.hdc
        End If
        DDC = frmBack.PicMainWork.hdc
        BitBlt DDC, 0, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY
        
        lRet = DoMorphingWire(TransPoints, FinishPoints, isQZsortWIRE, , Not isWireSingleColor, SpeedXangle, SpeedYangle, SpeedZangle)
        
        Blit3D
        If isSleep = True Then
            Sleep 16
        End If
        ShowFPS
        DoEvents
    Loop
    
    For i = 0 To frmMain.ImgPIX.Count - 1
        frmMain.ImgPIX(i).Enabled = True
    Next i
    
    
    For i = 0 To frmMain.ImgWIRE.Count - 1
        frmMain.ImgWIRE(i).Enabled = True
    Next i

    ReDim Points(StaticFinalPoints)
    For i = 0 To StaticFinalPoints 'UBound(FinishPoints)
        Points(i) = FinishPoints(i)
    Next i
    lRet = False
    Do Until lRet = True
        If isBackGround = True Then
            SDC = frmSprites.PicMainBackGround.hdc
        Else
            SDC = frmBack.PicMainClean.hdc
        End If
        DDC = frmBack.PicMainWork.hdc
        BitBlt DDC, 0, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY
        
        lRet = DoMorphingWire(TransPoints, FinishPoints, isQZsortWIRE, False, Not isWireSingleColor, SpeedXangle, SpeedYangle, SpeedZangle)

        Blit3D
        If isSleep = True Then
            Sleep 16
        End If
        ShowFPS
        DoEvents
    Loop
End Sub

Public Sub ApplyEngine(MeshFile As String)
Dim i As Integer
Dim hDCSPR As Long
    PrepareMorphing Points, FinishPoints, MeshFile
    lRet = False

    For i = 0 To frmMain.ImgPIX.Count - 1
        frmMain.ImgPIX(i).Enabled = False
    Next i
    
    For i = 0 To frmMain.ImgWIRE.Count - 1
        frmMain.ImgWIRE(i).Enabled = False
    Next i
    Do Until lRet = True
        If isBackGround = True Then
            SDC = frmSprites.PicMainBackGround.hdc
        Else
            SDC = frmBack.PicMainClean.hdc
        End If
        DDC = frmBack.PicMainWork.hdc
        BitBlt DDC, 0, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY
        
        If isBenny = True Then
            hDCSPR = frmSprites.PicSprBenny.hdc
        Else
           hDCSPR = frmSprites.PicSPR3d.hdc
        End If
        If isAlphaEffect = True And Not isPixelDOTT And Not isSolidSprite Then
            lRet = DoMorphing(TransPoints, FinishPoints, isQZsortDOTT, Not isPixelDOTT, hDCSPR, , Not isPixelSingleColorDOTT, isAlphaEffect, SpeedXangle, SpeedYangle, SpeedZangle, frmSprites.PicPlanets.hdc, True)
        Else
            lRet = DoMorphing(TransPoints, FinishPoints, isQZsortDOTT, Not isPixelDOTT, hDCSPR, , Not isPixelSingleColorDOTT, isAlphaEffect, SpeedXangle, SpeedYangle, SpeedZangle, , False)
        End If
        
        Blit3D
        If isSleep = True Then
            Sleep 16
        End If
        ShowFPS
        DoEvents
    Loop
    
    For i = 0 To frmMain.ImgPIX.Count - 1
        frmMain.ImgPIX(i).Enabled = True
    Next i
    
    For i = 0 To frmMain.ImgWIRE.Count - 1
        frmMain.ImgWIRE(i).Enabled = True
    Next i

    'Debug.Print "STOP"
    ReDim Points(StaticFinalPoints)
    
    For i = 0 To StaticFinalPoints
        Points(i) = FinishPoints(i)
    Next i
    lRet = False
    Do Until lRet = True
        If isBackGround = True Then
            SDC = frmSprites.PicMainBackGround.hdc
        Else
            SDC = frmBack.PicMainClean.hdc
        End If
        DDC = frmBack.PicMainWork.hdc
        BitBlt DDC, 0, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY
        
        If isAlphaEffect = True And Not isPixelDOTT And Not isSolidSprite Then
            DoSceneA Points, Faces, isBackGround, isQZsortDOTT, Not isPixelDOTT, hDCSPR, False, Not isPixelSingleColorDOTT, , SpeedXangle, SpeedYangle, SpeedZangle, frmSprites.PicPlanets.hdc, True, isAlphaEffect
        Else
            DoSceneA Points, Faces, isBackGround, isQZsortDOTT, Not isPixelDOTT, hDCSPR, False, Not isPixelSingleColorDOTT, , SpeedXangle, SpeedYangle, SpeedZangle, , , isAlphaEffect
        End If
        
        Blit3D
        If isSleep = True Then
            Sleep 16
        End If
        ShowFPS
        DoEvents
    Loop

End Sub
