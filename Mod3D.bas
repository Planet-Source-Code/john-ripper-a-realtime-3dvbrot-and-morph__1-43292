Attribute VB_Name = "Mod3D"
Option Explicit


Const XOrg      As Integer = 0          'Origin X of 3D rotations
Const YOrg      As Integer = 0
Const ZOrg      As Integer = 260

Const NumSinVal As Integer = 1024       'Values os precalc Sinus and Cosinus

Public NumPoints    As Integer
Public NumFaces     As Integer

Public XCenter As Long                  'Center of scene
Public YCenter As Long

Public XScreen As Integer               'Height and widht of main Picture Scene
Public YScreen As Integer

Public Type Point3D
    X As Long
    Y As Long
    Z As Long
    Aux As Long
End Type

Public Points() As Point3D
Public TempPoints() As Point3D
Public TransPoints() As Point3D
Public FinishPoints() As Point3D
Public StaticFinalPoints As Long

Public Type Face3D
    A As Long
    B As Long
    C As Long
    Z As Long
    AB As Long
    BC As Long
    CA As Long
End Type
Public Faces() As Face3D

Public Type CPalette
    Red As Integer
    Green As Integer
    Blue As Integer
End Type
Public MyPalette(255) As CPalette

Public CosTable(1025) As Long
Public SinTable(1025) As Long

Public Const PI = 3.141592654

Public Xangle As Integer
Public Yangle As Integer
Public Zangle As Integer

Public SpeedXangle As Long
Public SpeedYangle As Long
Public SpeedZangle As Long


Public Sub CreatePalette(Optional Red As Boolean = False, Optional Green As Boolean, Optional Blue As Boolean = False)
Dim i As Integer
    For i = 1 To 255
        If Red = True Then
            MyPalette(i).Red = i
        Else
            MyPalette(i).Red = 0
        End If
        If Green = True Then
            MyPalette(i).Green = i
        Else
            MyPalette(i).Green = 0
        End If
        If Blue = True Then
            MyPalette(i).Blue = i
        Else
            MyPalette(i).Blue = 0
        End If
    Next i
End Sub

Public Sub MakeCosTable()
    Dim CntVal As Long
    Dim CntAng As Single
    Dim IncDeg As Single
  
    IncDeg = 2 * PI / NumSinVal
    CntAng = IncDeg
    CntVal = 0
    Do Until CntVal > 1024
        CosTable(CntVal) = CInt((255 * Cos(CntAng)))
        CntAng = CntAng + IncDeg
        CntVal = CntVal + 1
    Loop
End Sub

Public Sub MakeSinTable()
    Dim CntVal As Long
    Dim CntAng As Single
    Dim IncDeg As Single

    IncDeg = 2 * PI / NumSinVal
    CntAng = IncDeg
    CntVal = 0
    Do Until CntVal > 1024
        SinTable(CntVal) = CInt((255 * Sin(CntAng)))
        CntAng = CntAng + IncDeg
        CntVal = CntVal + 1
    Loop
End Sub

Public Sub Calc3DRotations(SinX As Long, CosX As Long, SinY As Long, CosY As Long, SinZ As Long, CosZ As Long, _
                           OrgPoints() As Point3D, ByRef DesPoints() As Point3D, _
                           NumPoints As Long)

                                                              
  Dim x1        As Long
  Dim y1        As Long
  Dim Z1        As Long
  Dim Var32_1   As Long
  Dim Var32_2   As Long
  Dim CntPoints As Long

    For CntPoints = 0 To NumPoints '
        
'     X1 := (cos(YAngle) * X  - sin(YAngle) * Z)
        Var32_1 = CLng((CosY * OrgPoints(CntPoints).X) / 256)
        Var32_2 = CLng((SinY * OrgPoints(CntPoints).Z) / 256)
        x1 = Var32_1 - Var32_2
'     Z1 := (sin(YAngle) * X  + cos(YAngle) * Z)
        Var32_1 = CLng((SinY * OrgPoints(CntPoints).X) / 256)
        Var32_2 = CLng((CosY * OrgPoints(CntPoints).Z) / 256)
        Z1 = Var32_1 + Var32_2
'     X  := (cos(ZAngle) * X1 + sin(ZAngle) * Y)
        Var32_1 = CLng((CosZ * x1) / 256)
        Var32_2 = CLng((SinZ * OrgPoints(CntPoints).Y) / 256)
        DesPoints(CntPoints).X = Var32_1 + Var32_2
'     Y1 := (cos(ZAngle) * Y  - sin(ZAngle) * X1)
        Var32_1 = CLng((CosZ * OrgPoints(CntPoints).Y) / 256)
        Var32_2 = CLng((SinZ * x1) / 256)
        y1 = Var32_1 - Var32_2
'     Z  := (cos(XAngle) * Z1 - sin(XAngle) * Y1)
        Var32_1 = CLng((CosX * Z1) / 256)
        Var32_2 = CLng((SinX * y1) / 256)
        DesPoints(CntPoints).Z = Var32_1 - Var32_2
'     Y  := (sin(XAngle)) * Z1 + cos(XAngle) * Y1)
        Var32_1 = CLng((SinX * Z1) / 256)
        Var32_2 = CLng((CosX * y1) / 256)
        DesPoints(CntPoints).Y = Var32_1 + Var32_2
      
        DesPoints(CntPoints).Aux = OrgPoints(CntPoints).Aux
   
   Next CntPoints

End Sub


'Return TRUE if face is visible
'if FaceVisible=False NOT PAINT anything. Increase speed!
Public Function FaceVisible(x1 As Long, y1 As Long, x2 As Long, y2 As Long, X3 As Long, Y3 As Long) As Boolean
'simple escalar product
Dim A As Long
Dim B As Long
    If isHidenFacesWIRE = False Then
        FaceVisible = True
        Exit Function
    End If
    FaceVisible = True
    A = (x2 - x1) * (Y3 - y1)
    B = (X3 - x1) * (y2 - y1)
    If (A - B) < 0 Then
        FaceVisible = False
    End If
End Function


'Quick sort Subs to display first the nearest points

Public Sub QuickSortZPoints(NumPoints As Long, Points2qS() As Point3D, ByRef Points2qSDest() As Point3D)
    QuickSortPoints Points2qS, 0, NumPoints
End Sub
Private Sub QuickSortPoints(ByRef vntArr() As Point3D, _
    lngLeft As Long, lngRight As Long)

    Dim i As Long
    Dim j As Long
    Dim lngMid As Long
    Dim vntTestVal As Variant
    Dim vntTemp As Point3D
    
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        vntTestVal = vntArr(lngMid).Z
        i = lngLeft
        j = lngRight
        Do
            Do While vntArr(i).Z < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j).Z > vntTestVal
                j = j - 1
            Loop
            If i <= j Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j

        ' Optimize sort by sorting smaller segment first
        If j <= lngMid Then
            Call QuickSortPoints(vntArr, lngLeft, j)
            Call QuickSortPoints(vntArr, i, lngRight)
        Else
            Call QuickSortPoints(vntArr, i, lngRight)
            Call QuickSortPoints(vntArr, lngLeft, j)
        End If
    End If
End Sub

Public Sub QuickSortZFaces(NumPoints As Long, Points2qS() As Point3D, ByRef Faces2qS() As Face3D)
Dim cnt As Long
    For cnt = 0 To NumPoints
     
        Faces2qS(cnt).Z = (Points2qS(Faces2qS(cnt).A).Z + _
                      Points2qS(Faces2qS(cnt).B).Z + _
                      Points2qS(Faces2qS(cnt).C).Z) / 3
    Next cnt
' { reOrder it QUICKLY!!!! }
   QuickSortFaces Faces2qS, 0, NumPoints
End Sub
Private Sub QuickSortFaces(ByRef vntArr() As Face3D, _
    lngLeft As Long, lngRight As Long)

    Dim i As Long
    Dim j As Long
    Dim lngMid As Long
    Dim vntTestVal As Variant
    Dim vntTemp As Face3D
    
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        vntTestVal = vntArr(lngMid).Z
        i = lngLeft
        j = lngRight
        Do
            Do While vntArr(i).Z < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j).Z > vntTestVal
                j = j - 1
            Loop
            If i <= j Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j

        ' Optimize sort by sorting smaller segment first
        If j <= lngMid Then
            Call QuickSortFaces(vntArr, lngLeft, j)
            Call QuickSortFaces(vntArr, i, lngRight)
        Else
            Call QuickSortFaces(vntArr, i, lngRight)
            Call QuickSortFaces(vntArr, lngLeft, j)
        End If
    End If
End Sub


Public Sub Proyect3D(XScreen As Long, YScreen As Long, NumPoints As Long, _
                   OrgPoints() As Point3D, ByRef DesPoints() As Point3D)

    Dim CntPoints As Long

    For CntPoints = 0 To NumPoints
        If OrgPoints(CntPoints).Z >= ZOrg Then
            DesPoints(CntPoints).X = 320
            DesPoints(CntPoints).Y = 200
        Else
          DesPoints(CntPoints).X = XScreen + ((XOrg * OrgPoints(CntPoints).Z - OrgPoints(CntPoints).X * ZOrg) / (OrgPoints(CntPoints).Z - ZOrg))
          DesPoints(CntPoints).Y = YScreen + ((YOrg * OrgPoints(CntPoints).Z - OrgPoints(CntPoints).Y * ZOrg) / (OrgPoints(CntPoints).Z - ZOrg))
        End If
    Next CntPoints

End Sub

Public Sub PrepareMorphing(ByRef OriginalPoints() As Point3D, ByRef FinalPoints() As Point3D, FileMeshDestination As String, Optional ReadFaces As Boolean = False)
Dim i As Integer
Dim rndPoint As Long
    
    ReadFileMesh FileMeshDestination, FinalPoints, Faces, ReadFaces
    ReDim TransPoints(UBound(FinalPoints))
    StaticFinalPoints = UBound(FinalPoints)
    If UBound(OriginalPoints) = UBound(FinalPoints) Then
        For i = 0 To UBound(OriginalPoints)
            TransPoints(i).X = OriginalPoints(i).X
            TransPoints(i).Y = OriginalPoints(i).Y
            TransPoints(i).Z = OriginalPoints(i).Z
            TransPoints(i).Aux = FinalPoints(i).Aux
        Next i
    End If
    
    
    If UBound(OriginalPoints) < UBound(FinalPoints) Then
        For i = 0 To UBound(OriginalPoints)
            TransPoints(i).X = OriginalPoints(i).X
            TransPoints(i).Y = OriginalPoints(i).Y
            TransPoints(i).Z = OriginalPoints(i).Z
            TransPoints(i).Aux = FinalPoints(i).Aux
        Next i
        
        Randomize
        For i = UBound(OriginalPoints) + 1 To UBound(FinalPoints)
            rndPoint = CLng(UBound(OriginalPoints) * Rnd)
            TransPoints(i).X = OriginalPoints(rndPoint).X
            TransPoints(i).Y = OriginalPoints(rndPoint).Y
            TransPoints(i).Z = OriginalPoints(rndPoint).Z
            TransPoints(i).Aux = FinalPoints(rndPoint).Aux
        Next i
    End If

    If UBound(OriginalPoints) > UBound(FinalPoints) Then
        ReDim Preserve FinalPoints(UBound(OriginalPoints))
        ReDim TransPoints(UBound(OriginalPoints))
        ReDim TempPoints(UBound(OriginalPoints))
        
        For i = 0 To UBound(OriginalPoints)
            TransPoints(i).X = OriginalPoints(i).X
            TransPoints(i).Y = OriginalPoints(i).Y
            TransPoints(i).Z = OriginalPoints(i).Z
            TransPoints(i).Aux = FinalPoints(i).Aux
        Next i
        
        Randomize
        For i = NumPoints + 1 To UBound(OriginalPoints)
            rndPoint = CLng(NumPoints * Rnd)
            FinalPoints(i).X = FinalPoints(rndPoint).X
            FinalPoints(i).Y = FinalPoints(rndPoint).Y
            FinalPoints(i).Z = FinalPoints(rndPoint).Z
            TransPoints(i).Aux = FinalPoints(rndPoint).Aux
        Next i
       
    End If
    
End Sub
Public Sub DoRotation(ByRef Points() As Point3D, Optional QSortZ As Boolean = False, Optional SPRITES As Boolean = False, Optional hDCSprites As Long, Optional FlagMorphing As Boolean, Optional ColorizePixel As Boolean = False, Optional AlphaEffect As Boolean = False, Optional incXangle As Long = 0, Optional incYangle As Long = 0, Optional incZangle As Long = 0, Optional hDCSpritesPlanets = 0, Optional PhsicodelicEffect As Boolean)
Dim lRet As Boolean
    lRet = DoMorphing(Points, Points, QSortZ, SPRITES, hDCSprites, False, ColorizePixel, AlphaEffect, incXangle, incYangle, incZangle, hDCSpritesPlanets, PhsicodelicEffect)
End Sub


Public Function DoMorphingWire(ByRef OriginalPoints() As Point3D, ByRef FinalPoints() As Point3D, Optional QSortZ As Boolean = False, Optional FlagMorphing As Boolean = True, Optional ColorizeWire As Boolean = False, Optional incXangle As Long = 0, Optional incYangle As Long = 0, Optional incZangle As Long = 0) As Boolean
Dim xx1 As Long
Dim xx2 As Long
Dim yy1 As Long
Dim yy2 As Long
Dim i As Integer
Dim WireColor As Integer
    If FlagMorphing = True Then
        Calc3DRotations SinTable(Xangle), CosTable(Xangle), SinTable(Yangle), CosTable(Yangle), SinTable(Zangle), CosTable(Zangle), TransPoints, TempPoints, UBound(FinalPoints)
    Else
        Calc3DRotations SinTable(Xangle), CosTable(Xangle), SinTable(Yangle), CosTable(Yangle), SinTable(Zangle), CosTable(Zangle), OriginalPoints, TempPoints, UBound(FinalPoints)
    End If
    
    If QSortZ = True Then
        QuickSortZFaces UBound(Faces), TempPoints, Faces
    End If
    
    Proyect3D XCenter, YCenter, UBound(FinalPoints), TempPoints, TempPoints
    
    For i = 0 To UBound(Faces)
        If ColorizeWire Then
            WireColor = TempPoints(Faces(i).A).Z + FogValueWIRE + 16
            If WireColor > 255 Then
                WireColor = 255
            End If
            If WireColor < 0 Then
                WireColor = 30
            End If
        Else
            Dim rVal As Integer
            Dim gVal As Integer
            Dim bVal As Integer
            If isRedPixelWIRE = True Then
                rVal = 255
            Else
                rVal = 0
            End If
            If isGreenPixelWIRE = True Then
                gVal = 255
            Else
                gVal = 0
            End If
            If isBluePixelWIRE = True Then
                bVal = 255
            Else
                bVal = 0
            End If
        End If
        If FaceVisible(TempPoints(Faces(i).A).X, TempPoints(Faces(i).A).Y, TempPoints(Faces(i).B).X, TempPoints(Faces(i).B).Y, TempPoints(Faces(i).C).X, TempPoints(Faces(i).C).Y) = True Then
    
            If isNoDiagonals = True Then
                If Faces(i).AB = 1 Then
                    xx1 = TempPoints(Faces(i).A).X
                    yy1 = TempPoints(Faces(i).A).Y
    
                    xx2 = TempPoints(Faces(i).B).X
                    yy2 = TempPoints(Faces(i).B).Y
    
                    frmBack.PicMainWork.Line (xx1, yy1)-(xx2, yy2)
                    If ColorizeWire = True Then
                        frmBack.PicMainWork.ForeColor = RGB(MyPalette(WireColor).Red, MyPalette(WireColor).Green, MyPalette(WireColor).Blue)
                    Else
                        frmBack.PicMainWork.ForeColor = RGB(rVal, gVal, bVal)
                    End If
                End If
            Else
                xx1 = TempPoints(Faces(i).A).X
                yy1 = TempPoints(Faces(i).A).Y
    
                xx2 = TempPoints(Faces(i).B).X
                yy2 = TempPoints(Faces(i).B).Y
    
                frmBack.PicMainWork.Line (xx1, yy1)-(xx2, yy2)
                If ColorizeWire = True Then
                    frmBack.PicMainWork.ForeColor = RGB(MyPalette(WireColor).Red, MyPalette(WireColor).Green, MyPalette(WireColor).Blue)
                Else
                    frmBack.PicMainWork.ForeColor = RGB(rVal, gVal, bVal)
                End If
                
            End If
            If isNoDiagonals = True Then
                If Faces(i).BC = 1 Then
                    xx1 = TempPoints(Faces(i).B).X
                    yy1 = TempPoints(Faces(i).B).Y
    
                    xx2 = TempPoints(Faces(i).C).X
                    yy2 = TempPoints(Faces(i).C).Y
    
                    frmBack.PicMainWork.Line (xx1, yy1)-(xx2, yy2)
                    If ColorizeWire = True Then
                        frmBack.PicMainWork.ForeColor = RGB(MyPalette(WireColor).Red, MyPalette(WireColor).Green, MyPalette(WireColor).Blue)
                    Else
                        frmBack.PicMainWork.ForeColor = RGB(rVal, gVal, bVal)
                    End If
                    
                End If
            Else
                xx1 = TempPoints(Faces(i).B).X
                yy1 = TempPoints(Faces(i).B).Y
    
                xx2 = TempPoints(Faces(i).C).X
                yy2 = TempPoints(Faces(i).C).Y
    
                frmBack.PicMainWork.Line (xx1, yy1)-(xx2, yy2)
                If ColorizeWire = True Then
                    frmBack.PicMainWork.ForeColor = RGB(MyPalette(WireColor).Red, MyPalette(WireColor).Green, MyPalette(WireColor).Blue)
                Else
                    frmBack.PicMainWork.ForeColor = RGB(rVal, gVal, bVal)
                End If
            
            End If
            If isNoDiagonals = True Then
                If Faces(i).CA = 1 Then
                    xx1 = TempPoints(Faces(i).C).X
                    yy1 = TempPoints(Faces(i).C).Y
    
                    xx2 = TempPoints(Faces(i).A).X
                    yy2 = TempPoints(Faces(i).A).Y
    
                    frmBack.PicMainWork.Line (xx1, yy1)-(xx2, yy2)
                    If ColorizeWire = True Then
                        frmBack.PicMainWork.ForeColor = RGB(MyPalette(WireColor).Red, MyPalette(WireColor).Green, MyPalette(WireColor).Blue)
                    Else
                        frmBack.PicMainWork.ForeColor = RGB(rVal, gVal, bVal)
                    End If
                
                End If
            Else
                xx1 = TempPoints(Faces(i).C).X
                yy1 = TempPoints(Faces(i).C).Y
    
                xx2 = TempPoints(Faces(i).A).X
                yy2 = TempPoints(Faces(i).A).Y
    
                frmBack.PicMainWork.Line (xx1, yy1)-(xx2, yy2)
                If ColorizeWire = True Then
                    frmBack.PicMainWork.ForeColor = RGB(MyPalette(WireColor).Red, MyPalette(WireColor).Green, MyPalette(WireColor).Blue)
                Else
                    frmBack.PicMainWork.ForeColor = RGB(rVal, gVal, bVal)
                End If
            
            End If

        End If
        
        
    Next i
    
    For i = 0 To UBound(FinalPoints)
        If FlagMorphing = True Then

            If TransPoints(i).X < FinalPoints(i).X Then
                TransPoints(i).X = TransPoints(i).X + 1
            End If
            
            If TransPoints(i).X > FinalPoints(i).X Then
                TransPoints(i).X = TransPoints(i).X - 1
            End If
            
            
            If TransPoints(i).Y < FinalPoints(i).Y Then
                TransPoints(i).Y = TransPoints(i).Y + 1
            End If
            
            If TransPoints(i).Y > FinalPoints(i).Y Then
                TransPoints(i).Y = TransPoints(i).Y - 1
            End If
            
            
            If TransPoints(i).Z < FinalPoints(i).Z Then
                TransPoints(i).Z = TransPoints(i).Z + 1
            End If
            
            If TransPoints(i).Z > FinalPoints(i).Z Then
                TransPoints(i).Z = TransPoints(i).Z - 1
            End If
        End If
    Next i
    
    If FlagMorphing = True Then
        DoMorphingWire = True
        For i = 0 To UBound(FinalPoints)

            If TransPoints(i).X <> FinalPoints(i).X Or _
                TransPoints(i).Y <> FinalPoints(i).Y Or _
                TransPoints(i).Z <> FinalPoints(i).Z Then
                DoMorphingWire = False
                Exit For
            End If
        Next i
    End If
    
    Xangle = Xangle + incXangle '0 '2
    If Xangle > 1024 Then
        Xangle = Xangle - 1024
    End If
    If Xangle < 0 Then
        Xangle = 0
    End If
    Yangle = Yangle + incYangle '3
    If Yangle > 1024 Then
        Yangle = Yangle - 1024
    End If
    If Yangle < 0 Then
        Yangle = 0
    End If
    
    Zangle = Zangle + incZangle '0 '3
    If Zangle > 1024 Then
        Zangle = Zangle - 1024
    End If
    If Zangle < 0 Then
        Zangle = 0
    End If
End Function

Public Function DoMorphing(ByRef OriginalPoints() As Point3D, ByRef FinalPoints() As Point3D, Optional QSortZ As Boolean = False, Optional SPRITES As Boolean = False, Optional hDCSprites As Long, Optional FlagMorphing As Boolean = True, Optional ColorizePixel As Boolean = False, Optional AlphaEffect As Boolean = False, Optional incXangle As Long = 0, Optional incYangle As Long = 0, Optional incZangle As Long = 0, Optional hDCSpritesPlanets = 0, Optional PhsicodelicEffect As Boolean) As Boolean
Dim i As Integer
Dim ZColorSPR(5) As Long '(0):Zmin, (5):Zmax
Dim ZetasSPRinc As Long
Dim SPRNumber As Long
Dim PixelColor As Long
Dim BlitMaskBenny As Boolean

    If FlagMorphing = True Then
        Calc3DRotations SinTable(Xangle), CosTable(Xangle), SinTable(Yangle), CosTable(Yangle), SinTable(Zangle), CosTable(Zangle), TransPoints, TempPoints, UBound(FinalPoints)
    Else
        Calc3DRotations SinTable(Xangle), CosTable(Xangle), SinTable(Yangle), CosTable(Yangle), SinTable(Zangle), CosTable(Zangle), OriginalPoints, TempPoints, UBound(FinalPoints)
    End If
    If QSortZ = True Then
        QuickSortZPoints UBound(FinalPoints), TempPoints, TempPoints
    End If
        
    
    Proyect3D XCenter, YCenter, UBound(FinalPoints), TempPoints, TempPoints
    
    If SPRITES = True Then
        SDC = hDCSprites
        DDC = frmBack.PicMainWork.hdc
    
        If isBenny = False Then
            ZColorSPR(0) = TempPoints(0).Z
            ZColorSPR(5) = TempPoints(UBound(Points)).Z
            ZetasSPRinc = (Abs(ZColorSPR(0)) + Abs(ZColorSPR(5))) \ 5
            For i = 1 To 4
                ZColorSPR(i) = ZColorSPR(i - 1) + ZetasSPRinc
            Next i
        End If
    Else
        SDC = frmBack.PicMainWork.hdc
    End If
    

    For i = 0 To UBound(FinalPoints)

        If SPRITES = False Then
            If ColorizePixel Then
                PixelColor = (TempPoints(i).Z) + FogValueDOTT + 32
                If PixelColor > 255 Then
                    PixelColor = 255
                End If
                If PixelColor < 0 Then
                    PixelColor = 30
                End If
                    
                SetPixelV DDC, TempPoints(i).X, TempPoints(i).Y, RGB(MyPalette(PixelColor).Red, MyPalette(PixelColor).Green, MyPalette(PixelColor).Blue)
            Else
                Dim rVal As Integer
                Dim gVal As Integer
                Dim bVal As Integer
                If isRedPixelDOTT = True Then
                    rVal = 255
                Else
                    rVal = 0
                End If
                If isGreenPixelDOTT = True Then
                    gVal = 255
                Else
                    gVal = 0
                End If
                If isBluePixelDOTT = True Then
                    bVal = 255
                Else
                    bVal = 0
                End If
                
                SetPixelV DDC, TempPoints(i).X, TempPoints(i).Y, RGB(rVal, gVal, bVal)
            End If
        Else
            If isBenny = False Then
                Select Case TempPoints(i).Z
                    Case ZColorSPR(0), ZColorSPR(1)
                        SPRNumber = 4
                    Case ZColorSPR(1), ZColorSPR(2)
                        SPRNumber = 3
                    Case ZColorSPR(2), ZColorSPR(3)
                        SPRNumber = 2
                    Case ZColorSPR(3), ZColorSPR(4)
                        SPRNumber = 1
                    Case ZColorSPR(4), ZColorSPR(5)
                        SPRNumber = 0
                End Select
                BlitMaskBenny = True
            Else
                If TempPoints(i).Aux <> 16711935 Then
                    Select Case TempPoints(i).Aux
                        Case 16776447   'blanco
                            SPRNumber = 0
                        Case 16776344   'Azul + claro
                            SPRNumber = 1
                        Case 1677619
                            SPRNumber = 2
                        Case 13683712
                            SPRNumber = 3
                        Case 11580416
                            SPRNumber = 4
                        Case 8420352
                            SPRNumber = 5
                        Case 6841344
                            SPRNumber = 6
                        Case 0
                            SPRNumber = 7
                    End Select
                    BlitMaskBenny = True
                Else
                    BlitMaskBenny = False
                End If
            End If
            If AlphaEffect = True Then
            Else
                If hDCSpritesPlanets = 0 Then
                    If BlitMaskBenny = True Then
                        BitBlt DDC, TempPoints(i).X, TempPoints(i).Y, 8, 8, SDC, 0, 8, SRCAND
                    End If
                Else
                    BitBlt DDC, TempPoints(i).X, TempPoints(i).Y, MyPlanet(TempPoints(i).Aux).Width, MyPlanet(TempPoints(i).Aux).Height, hDCSpritesPlanets, MyPlanet(TempPoints(i).Aux).PosX, 58, SRCAND
                End If
            End If
            If hDCSpritesPlanets = 0 Then
                If isShadowedSprites = True Then
                    If BlitMaskBenny = True Then
                        BitBlt DDC, TempPoints(i).X, TempPoints(i).Y, 8, 8, SDC, 0 + (8 * SPRNumber), 0, SRCPAINT
                    End If
                Else
                    BitBlt DDC, TempPoints(i).X, TempPoints(i).Y, 8, 8, SDC, 0, 0, SRCPAINT
                End If
            Else
                If isBenny = False Then
                    If PhsicodelicEffect = False Then
                        BitBlt DDC, TempPoints(i).X, TempPoints(i).Y, MyPlanet(TempPoints(i).Aux).Width, MyPlanet(TempPoints(i).Aux).Height, hDCSpritesPlanets, MyPlanet(TempPoints(i).Aux).PosX, 0, SRCPAINT
                    Else
                        BitBlt DDC, TempPoints(i).X, TempPoints(i).Y, MyPlanet(TempPoints(i).Aux).Width, MyPlanet(TempPoints(i).Aux).Height, SDC, MyPlanet(TempPoints(i).Aux).PosX, 0, SRCPAINT
                    End If
                End If
            End If
        End If

        If FlagMorphing = True Then
            
            If TransPoints(i).X < FinalPoints(i).X Then
                TransPoints(i).X = TransPoints(i).X + 1
            End If
            
            If TransPoints(i).X > FinalPoints(i).X Then
                TransPoints(i).X = TransPoints(i).X - 1
            End If
            
            
            If TransPoints(i).Y < FinalPoints(i).Y Then
                TransPoints(i).Y = TransPoints(i).Y + 1
            End If
            
            If TransPoints(i).Y > FinalPoints(i).Y Then
                TransPoints(i).Y = TransPoints(i).Y - 1
            End If
            
            
            If TransPoints(i).Z < FinalPoints(i).Z Then
                TransPoints(i).Z = TransPoints(i).Z + 1
            End If
            
            If TransPoints(i).Z > FinalPoints(i).Z Then
                TransPoints(i).Z = TransPoints(i).Z - 1
            End If
        End If
    Next i
        
    If FlagMorphing = True Then
        DoMorphing = True
        For i = 0 To UBound(FinalPoints)
            If TransPoints(i).X <> FinalPoints(i).X Or _
                TransPoints(i).Y <> FinalPoints(i).Y Or _
                TransPoints(i).Z <> FinalPoints(i).Z Then
                DoMorphing = False
                Exit For
            End If
        Next i
    End If
    Xangle = Xangle + incXangle '0 '2
    If Xangle > 1024 Then
        Xangle = Xangle - 1024
    End If
    If Xangle < 0 Then
        Xangle = 0
    End If
    Yangle = Yangle + incYangle '3
    If Yangle > 1024 Then
        Yangle = Yangle - 1024
    End If
    If Yangle < 0 Then
        Yangle = 0
    End If
    
    Zangle = Zangle + incZangle '0 '3
    If Zangle > 1024 Then
        Zangle = Zangle - 1024
    End If
    If Zangle < 0 Then
        Zangle = 0
    End If

End Function

Public Sub Blit3D()
    SDC = frmBack.PicMainWork.hdc
    DDC = frmMain.PicMain.hdc
    BitBlt DDC, 0, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY
End Sub

