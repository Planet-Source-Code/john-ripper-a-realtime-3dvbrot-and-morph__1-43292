Attribute VB_Name = "ModMisc"
Option Explicit

Public SDC As Long  'Source hDC for Blitting
Public DDC As Long  'Destination hDC for Blitting

Public CounterSeconds As Long

Public rHeight As Long
Public rWidth As Long

Public rHeightASC As Long
Public rWidthASC As Long

Public Type tPlanet
    PosX As Integer
    Width As Integer
    Height As Integer
End Type

Public MyPlanet(8) As tPlanet

Public WhatPIXrunning As Integer
Public WhatWIRErunning As Integer

Public Sub InitializePlanets()
'Load info for the srpites of planets (see frmSprites.PicPlanetes)
    MyPlanet(0).PosX = 1
    MyPlanet(0).Width = 14
    MyPlanet(0).Height = 14
    
    MyPlanet(1).PosX = 16
    MyPlanet(1).Width = 26
    MyPlanet(1).Height = 26
    
    MyPlanet(2).PosX = 43
    MyPlanet(2).Width = 21
    MyPlanet(2).Height = 22

    MyPlanet(3).PosX = 65
    MyPlanet(3).Width = 19
    MyPlanet(3).Height = 20

    MyPlanet(4).PosX = 85
    MyPlanet(4).Width = 57
    MyPlanet(4).Height = 57

    MyPlanet(5).PosX = 143
    MyPlanet(5).Width = 58
    MyPlanet(5).Height = 43

    MyPlanet(6).PosX = 202
    MyPlanet(6).Width = 31
    MyPlanet(6).Height = 31

    MyPlanet(7).PosX = 234
    MyPlanet(7).Width = 25
    MyPlanet(7).Height = 26

    MyPlanet(8).PosX = 260
    MyPlanet(8).Width = 9
    MyPlanet(8).Height = 9

End Sub

'This Sub Read the MeshFile and stores data on ByRef Params. Open any dVB file with NotePad to see internal strucutre)
Public Sub ReadFileMesh(FileMesh As String, ByRef arrPoints() As Point3D, ByRef arrFaces() As Face3D, Optional ReadFaces As Boolean = False)
    Dim dataFile As String
    Dim nF As Integer
    Dim i As Integer
    Dim FilePoints As Long
    Dim FileFaces As Long
    Dim FlagFaces As Boolean
    Dim CounterFile As Long
    Dim Pos1 As Long
    Dim Pos2 As Long
    Dim Pos3 As Long
    Dim Pos4 As Long
    Dim Pos5 As Long

    nF = FreeFile
    Open FileMesh For Input As #nF
    For i = 1 To 8          'read "header"
        Line Input #nF, dataFile
    Next i

    Line Input #nF, dataFile
    Pos1 = InStr(1, dataFile, "=")
    FilePoints = Mid(dataFile, Pos1 + 1)
    NumPoints = FilePoints
    ReDim arrPoints(FilePoints)
    ReDim TempPoints(FilePoints)
    
    Line Input #nF, dataFile
    FlagFaces = True
    If InStr(1, dataFile, "Not Available") <> 0 Then
        FlagFaces = False
    Else
        Pos1 = InStr(1, dataFile, "=")
        FileFaces = Mid(dataFile, Pos1 + 1)
        NumFaces = FileFaces
    End If
    
    Line Input #nF, dataFile        '""
    Line Input #nF, dataFile        '"--------------------------POINTS-------------------------"
    
    CounterFile = 0
    Do Until CounterFile = FilePoints + 1
        Line Input #nF, dataFile    'X!Y@Z format
        Pos1 = InStr(1, dataFile, "!")
        Pos2 = InStr(1, dataFile, "@")
        Pos3 = InStr(1, dataFile, "*")
        
        arrPoints(CounterFile).X = Mid(dataFile, 1, Pos1 - 1)
        arrPoints(CounterFile).Y = Mid(dataFile, Pos1 + 1, Pos2 - Pos1 - 1)
        If Pos3 = 0 Then
            arrPoints(CounterFile).Z = Mid(dataFile, Pos2 + 1)
        Else
            arrPoints(CounterFile).Z = Mid(dataFile, Pos2 + 1, Pos3 - Pos2 - 1)
            arrPoints(CounterFile).Aux = Mid(dataFile, Pos3 + 1)
        End If
        CounterFile = CounterFile + 1
    Loop
    
    If ReadFaces = True And FlagFaces = True Then
        ReDim arrFaces(FileFaces)
        
        Line Input #nF, dataFile    '--------------------------FACES--------------------------
    
        CounterFile = 0
        Do Until CounterFile = FileFaces + 1
            Line Input #nF, dataFile    'A!B@C format
            Pos1 = InStr(1, dataFile, "!")
            Pos2 = InStr(1, dataFile, "@")
            Pos3 = InStr(1, dataFile, "*")
            Pos4 = InStr(1, dataFile, "%")
            Pos5 = InStr(1, dataFile, "(")
            
            arrFaces(CounterFile).A = Mid(dataFile, 1, Pos1 - 1)
            arrFaces(CounterFile).B = Mid(dataFile, Pos1 + 1, Pos2 - Pos1 - 1)
            arrFaces(CounterFile).C = Mid(dataFile, Pos2 + 1, Pos3 - Pos2 - 1)
            arrFaces(CounterFile).Z = 0
            arrFaces(CounterFile).AB = Mid(dataFile, Pos3 + 1, Pos4 - Pos3 - 1)
            arrFaces(CounterFile).BC = Mid(dataFile, Pos4 + 1, Pos5 - Pos4 - 1)
            arrFaces(CounterFile).CA = Mid(dataFile, Pos5 + 1)
            CounterFile = CounterFile + 1
        Loop
    
    End If
    
    Close #nF
End Sub
Public Sub DoSceneA(arrPoints() As Point3D, arrFaces() As Face3D, Optional BackGround As Boolean = False, Optional QSortZ As Boolean = False, Optional SPRITES As Boolean = False, Optional hDCSprites As Long, Optional FlagMorphing As Boolean = False, Optional ColorizePixel As Boolean = False, Optional ScrollingBackGround As Boolean = False, Optional incXangle As Long = 0, Optional incYangle As Long = 0, Optional incZangle As Long = 0, Optional hDCSpritesPlanets = 0, Optional PhsicodelicEffect As Boolean, Optional AlphaEffect As Boolean)

    If ScrollingBackGround = False Then
        If BackGround = True Then
            SDC = frmSprites.PicMainBackGround.hdc
            DDC = frmBack.PicMainWork.hdc
            BitBlt DDC, 0, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY
        Else
            SDC = frmBack.PicMainClean.hdc
            DDC = frmBack.PicMainWork.hdc
            BitBlt DDC, 0, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY
        End If
    End If
    DoRotation arrPoints, QSortZ, SPRITES, hDCSprites, FlagMorphing, ColorizePixel, AlphaEffect, incXangle, incYangle, incZangle, hDCSpritesPlanets, PhsicodelicEffect
End Sub

'Not used in anytime.For futured uses
Public Sub ScrollBackGround(xCoord As Long)
    SDC = frmSprites.PicMainBackGround.hdc
    DDC = frmBack.PicMainWork.hdc
    BitBlt DDC, 0, 0, XScreen, YScreen, SDC, xCoord, 0, SRCCOPY
    
    SDC = frmBack.PicMainClean.hdc
    DDC = frmBack.PicMainWork.hdc
    BitBlt DDC, XScreen - xCoord, 0, XScreen, YScreen, SDC, 0, 0, SRCCOPY

End Sub

