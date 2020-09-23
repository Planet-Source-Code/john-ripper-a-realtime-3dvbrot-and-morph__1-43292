Attribute VB_Name = "ModASCII"
Option Explicit


Public Type tAscii
    CodeASC As Integer
    PosX As Integer
    Width As Integer
End Type

Public XaxisHor As Long
Public YaxisHor As Long
Public XaxisVer As Long
Public YaxisVer As Long

Public CounterMsgHor As Long

Public MyAscii(45) As tAscii
Public MsgHor(10) As String

Public Sub InitializeAscii()

    MyAscii(1).CodeASC = Asc(" ")
    MyAscii(1).PosX = 0
    MyAscii(1).Width = 24
    
    MyAscii(2).CodeASC = Asc("A")
    MyAscii(2).PosX = 25
    MyAscii(2).Width = 32
    
    MyAscii(3).CodeASC = Asc("B")
    MyAscii(3).PosX = 58
    MyAscii(3).Width = 30
    
    MyAscii(4).CodeASC = Asc("C")
    MyAscii(4).PosX = 89
    MyAscii(4).Width = 31
    
    MyAscii(5).CodeASC = Asc("D")
    MyAscii(5).PosX = 121
    MyAscii(5).Width = 32
    
    MyAscii(6).CodeASC = Asc("E")
    MyAscii(6).PosX = 154
    MyAscii(6).Width = 30
    
    MyAscii(7).CodeASC = Asc("F")
    MyAscii(7).PosX = 185
    MyAscii(7).Width = 30
    
    MyAscii(8).CodeASC = Asc("G")
    MyAscii(8).PosX = 216
    MyAscii(8).Width = 36
    
    MyAscii(9).CodeASC = Asc("H")
    MyAscii(9).PosX = 253
    MyAscii(9).Width = 32
    
    MyAscii(10).CodeASC = Asc("I")
    MyAscii(10).PosX = 286
    MyAscii(10).Width = 14
    
    MyAscii(11).CodeASC = Asc("J")
    MyAscii(11).PosX = 301
    MyAscii(11).Width = 17
    
    MyAscii(12).CodeASC = Asc("K")
    MyAscii(12).PosX = 319
    MyAscii(12).Width = 34
    
    MyAscii(13).CodeASC = Asc("L")
    MyAscii(13).PosX = 354
    MyAscii(13).Width = 24
    
    MyAscii(14).CodeASC = Asc("M")
    MyAscii(14).PosX = 379
    MyAscii(14).Width = 38
    
    MyAscii(15).CodeASC = Asc("N")
    MyAscii(15).PosX = 418
    MyAscii(15).Width = 32
    
    MyAscii(16).CodeASC = Asc("Ñ")
    MyAscii(16).PosX = 451
    MyAscii(16).Width = 24
    
    MyAscii(17).CodeASC = Asc("O")
    MyAscii(17).PosX = 476
    MyAscii(17).Width = 35
    
    MyAscii(18).CodeASC = Asc("P")
    MyAscii(18).PosX = 512
    MyAscii(18).Width = 31
    
    MyAscii(19).CodeASC = Asc("Q")
    MyAscii(19).PosX = 544
    MyAscii(19).Width = 35
    
    MyAscii(20).CodeASC = Asc("R")
    MyAscii(20).PosX = 580
    MyAscii(20).Width = 30
    
    MyAscii(21).CodeASC = Asc("S")
    MyAscii(21).PosX = 611
    MyAscii(21).Width = 32
    
    MyAscii(22).CodeASC = Asc("T")
    MyAscii(22).PosX = 644
    MyAscii(22).Width = 30
    
    MyAscii(23).CodeASC = Asc("U")
    MyAscii(23).PosX = 675
    MyAscii(23).Width = 29
    
    MyAscii(24).CodeASC = Asc("V")
    MyAscii(24).PosX = 705
    MyAscii(24).Width = 33
    
    MyAscii(25).CodeASC = Asc("W")
    MyAscii(25).PosX = 739
    MyAscii(25).Width = 46
    
    MyAscii(26).CodeASC = Asc("X")
    MyAscii(26).PosX = 786
    MyAscii(26).Width = 34
    
    MyAscii(27).CodeASC = Asc("Y")
    MyAscii(27).PosX = 821
    MyAscii(27).Width = 34
    
    MyAscii(28).CodeASC = Asc("Z")
    MyAscii(28).PosX = 856
    MyAscii(28).Width = 33
    
    MyAscii(29).CodeASC = Asc("!")
    MyAscii(29).PosX = 890
    MyAscii(29).Width = 12
    
    MyAscii(30).CodeASC = Asc("$")
    MyAscii(30).PosX = 903
    MyAscii(30).Width = 22
    
    MyAscii(31).CodeASC = Asc("(")
    MyAscii(31).PosX = 926
    MyAscii(31).Width = 12
    
    MyAscii(32).CodeASC = Asc(")")
    MyAscii(32).PosX = 939
    MyAscii(32).Width = 13
    
    MyAscii(33).CodeASC = Asc("?")
    MyAscii(33).PosX = 953
    MyAscii(33).Width = 26
    
    MyAscii(34).CodeASC = Asc("-")
    MyAscii(34).PosX = 980
    MyAscii(34).Width = 16
    
    MyAscii(35).CodeASC = Asc(".")
    MyAscii(35).PosX = 997
    MyAscii(35).Width = 12
            
    MyAscii(36).CodeASC = Asc("1")
    MyAscii(36).PosX = 0
    MyAscii(36).Width = 18
    
    MyAscii(37).CodeASC = Asc("2")
    MyAscii(37).PosX = 19
    MyAscii(37).Width = 30
    
    MyAscii(38).CodeASC = Asc("3")
    MyAscii(38).PosX = 50
    MyAscii(38).Width = 28
    
    MyAscii(39).CodeASC = Asc("4")
    MyAscii(39).PosX = 79
    MyAscii(39).Width = 33
    
    MyAscii(40).CodeASC = Asc("5")
    MyAscii(40).PosX = 113
    MyAscii(40).Width = 33
    
    MyAscii(41).CodeASC = Asc("6")
    MyAscii(41).PosX = 147
    MyAscii(41).Width = 27
    
    MyAscii(42).CodeASC = Asc("7")
    MyAscii(42).PosX = 175
    MyAscii(42).Width = 32
    
    MyAscii(43).CodeASC = Asc("8")
    MyAscii(43).PosX = 208
    MyAscii(43).Width = 27
    
    MyAscii(44).CodeASC = Asc("9")
    MyAscii(44).PosX = 236
    MyAscii(44).Width = 28
    
    MyAscii(45).CodeASC = Asc("0")
    MyAscii(45).PosX = 265
    MyAscii(45).Width = 35
        
End Sub

Public Sub WriteMyAscii(hDCOrg As Long, hDCDest As Long, DestWidth As Long, DisplayText As String, Xdest As Long, Ydest As Long, ByRef FinishScroll As Boolean, Optional Leading As Integer = 0, Optional Scrolling As Boolean = False)
Dim i As Integer
Dim j As Integer
Dim CounterX As Long
Dim UnknowAscii As Boolean
Dim OnlyChar As String
Dim tempAsc As tAscii
Dim TempUcase As String
    

    
    If Len(DisplayText) = 0 Then
        Exit Sub
    End If

    TempUcase = UCase$(DisplayText)
    CounterX = Xdest
        
    For i = 1 To Len(TempUcase)
        OnlyChar = Mid(TempUcase, i, 1)
        UnknowAscii = True
        For j = 1 To UBound(MyAscii)
            If Asc(OnlyChar) = MyAscii(j).CodeASC Then
                tempAsc = MyAscii(j)
                UnknowAscii = False
                Exit For
            End If
        Next j
    
        'only prints necessary Text
        If CounterX > DestWidth Or CounterX < -38 Then
            If Scrolling = False Then
                Exit For
            Else
                CounterX = CounterX + tempAsc.Width - Leading
            End If
        Else
            If i = Len(TempUcase) And CounterX < -38 - Leading Then
                FinishScroll = True
            Else
                FinishScroll = False
            End If
            If UnknowAscii = False Then
                If tempAsc.CodeASC >= Asc("0") And tempAsc.CodeASC <= Asc("9") Then 'numbers (2nd line of sprites)
                    BitBlt hDCDest, CounterX, Ydest, tempAsc.Width, 42, hDCOrg, tempAsc.PosX, 130, SRCAND
                    BitBlt hDCDest, CounterX, Ydest, tempAsc.Width, 42, hDCOrg, tempAsc.PosX, 87, SRCPAINT
                Else                    'letters and simbols (1st line of sprites)
                    BitBlt hDCDest, CounterX, Ydest, tempAsc.Width, 42, hDCOrg, tempAsc.PosX, 44, SRCAND
                    BitBlt hDCDest, CounterX, Ydest, tempAsc.Width, 42, hDCOrg, tempAsc.PosX, 1, SRCPAINT
                End If
                CounterX = CounterX + tempAsc.Width - Leading
            Else    'not valid asc-> print "?"
                BitBlt hDCDest, CounterX, Ydest, 26, 42, hDCOrg, 953, 44, SRCAND
                BitBlt hDCDest, CounterX, Ydest, 26, 42, hDCOrg, 953, 1, SRCPAINT
                CounterX = CounterX + 26 - Leading
            End If
        End If
        If i = Len(TempUcase) And CounterX < -38 - Leading Then
            FinishScroll = True
        Else
            FinishScroll = False
        End If
    Next i
    
End Sub
Public Sub CleanHorizontalAsc()
    SDC = frmBack.PicASCHorizontalClean.hdc
    DDC = frmBack.PicASCHorizontalWork.hdc
    BitBlt DDC, 0, 0, rWidthASC, rHeightASC, SDC, 0, 0, SRCCOPY
End Sub

Public Sub BlitHorizontalAsc()
    SDC = frmBack.PicASCHorizontalWork.hdc
    DDC = frmMain.PicASCHorizontal.hdc
    BitBlt DDC, 0, 0, rWidthASC, rHeightASC, SDC, 0, 0, SRCCOPY
End Sub

