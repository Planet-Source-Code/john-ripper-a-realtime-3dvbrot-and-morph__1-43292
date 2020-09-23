VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF0000&
   Caption         =   "3D Rotations and Real Time Morphing in pure VB SourceCode"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   594
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   796
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicASCHorizontal 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   7500
      ScaleHeight     =   30
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   128.03
      TabIndex        =   44
      ToolTipText     =   "Frames per second"
      Top             =   5460
      Width           =   1950
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4740
      Left            =   60
      ScaleHeight     =   316
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   486
      TabIndex        =   43
      Top             =   60
      Width           =   7290
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Misc."
      Height          =   855
      Left            =   3660
      TabIndex        =   39
      Top             =   7860
      Width           =   3675
      Begin VB.CheckBox CheckBackGround 
         BackColor       =   &H00FFFFC0&
         Caption         =   "BackGround Enable"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   300
         Width           =   1755
      End
      Begin VB.CheckBox CheckGraphFPS 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Graphical FPS Text"
         Height          =   195
         Left            =   1920
         TabIndex        =   41
         Top             =   300
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox CheckSleep 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enable Sleep API"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Enable 16 milisecons sleep API. Recomended for FASTER PCs"
         Top             =   600
         Width           =   1755
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rotation control"
      Height          =   1215
      Left            =   3660
      TabIndex        =   29
      Top             =   6600
      Width           =   3675
      Begin VB.HScrollBar HSx 
         Height          =   255
         Left            =   1320
         Max             =   10
         TabIndex        =   32
         Top             =   240
         Value           =   5
         Width           =   1755
      End
      Begin VB.HScrollBar HSy 
         Height          =   255
         Left            =   1320
         Max             =   10
         TabIndex        =   31
         Top             =   540
         Value           =   5
         Width           =   1755
      End
      Begin VB.HScrollBar HSz 
         Height          =   255
         Left            =   1320
         Max             =   10
         TabIndex        =   30
         Top             =   840
         Value           =   5
         Width           =   1755
      End
      Begin VB.Label lblXaxis 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         Height          =   255
         Left            =   3180
         TabIndex        =   38
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblYaxis 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         Height          =   255
         Left            =   3180
         TabIndex        =   37
         Top             =   540
         Width           =   375
      End
      Begin VB.Label lblZaxis 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         Height          =   255
         Left            =   3180
         TabIndex        =   36
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "X axis speed ->"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Y axis speed ->"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Z axis speed ->"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   900
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "WireFrame Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   3660
      TabIndex        =   18
      Top             =   4860
      Width           =   3675
      Begin VB.HScrollBar HSWidth 
         Height          =   195
         Left            =   2100
         Max             =   4
         Min             =   1
         TabIndex        =   49
         Top             =   540
         Value           =   1
         Width           =   795
      End
      Begin VB.CheckBox CheckDiagonals 
         BackColor       =   &H00008000&
         Caption         =   "No Diagonals"
         Height          =   195
         Left            =   2100
         TabIndex        =   48
         Top             =   300
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.OptionButton OptLines1Color 
         BackColor       =   &H00008000&
         Caption         =   "Single Color"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton OptLineXColors 
         BackColor       =   &H00008000&
         Caption         =   "Multi Color"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0000C000&
         Height          =   795
         Left            =   1380
         TabIndex        =   21
         Top             =   780
         Width           =   1575
         Begin VB.CheckBox CheckBlueLIN 
            BackColor       =   &H0000C000&
            Caption         =   "B"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1080
            TabIndex        =   25
            Top             =   180
            Value           =   1  'Checked
            Width           =   435
         End
         Begin VB.CheckBox CheckGreenLIN 
            BackColor       =   &H0000C000&
            Caption         =   "G"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   180
            Value           =   1  'Checked
            Width           =   435
         End
         Begin VB.CheckBox CheckRedLIN 
            BackColor       =   &H0000C000&
            Caption         =   "R"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   180
            Value           =   1  'Checked
            Width           =   435
         End
         Begin VB.HScrollBar HSFogLin 
            Height          =   255
            Left            =   120
            Max             =   128
            TabIndex        =   22
            Top             =   480
            Value           =   64
            Width           =   915
         End
         Begin VB.Label lblFogLIN 
            BackColor       =   &H0000C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "64"
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.CheckBox CheckHideFaces 
         BackColor       =   &H00008000&
         Caption         =   "Hide non Visible Faces"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox CheckQZsortFaces 
         BackColor       =   &H00008000&
         Caption         =   "QZsortFaces"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.Label lblWireWidth 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Left            =   3000
         TabIndex        =   50
         ToolTipText     =   "DrawWith. Values >1 are VERY VERY slowly :("
         Top             =   540
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Dotted Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   60
      TabIndex        =   1
      Top             =   4860
      Width           =   3555
      Begin VB.CheckBox CheckQZsortPoints 
         BackColor       =   &H00008000&
         Caption         =   "QZsortPoints"
         Height          =   255
         Left            =   2100
         TabIndex        =   17
         ToolTipText     =   "Order Z axis for display first the nearest points. Looks better but slow down speed"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.Frame FrameSpr 
         BackColor       =   &H00008000&
         Caption         =   "Sprite Options"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   2100
         Width           =   3195
         Begin VB.CheckBox CheckShadowedSPR 
            BackColor       =   &H00008000&
            Caption         =   "Shadowed Sprites"
            Height          =   255
            Left            =   1500
            TabIndex        =   46
            ToolTipText     =   "Degradate sprite color on Z axis. Nearest more shinny, Far sprites more darkness"
            Top             =   300
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.OptionButton OptIMAX 
            BackColor       =   &H00008000&
            Caption         =   "IMAX 3D effect (requieres Alpha)"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   10
            ToolTipText     =   "Apply 3D effect. If you have got a 3D glasses, like IMAX glasses, you will can see REAL 3D object! xDDDD "
            Top             =   840
            Width           =   2895
         End
         Begin VB.OptionButton OptSolidSPR 
            BackColor       =   &H00008000&
            Caption         =   "Solid Sprite"
            Height          =   195
            Left            =   180
            TabIndex        =   9
            ToolTipText     =   "Display ""Normal"" sprites"
            Top             =   600
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.CheckBox CheckAlphaSPR 
            BackColor       =   &H00008000&
            Caption         =   "Alpha Effect"
            Height          =   255
            Left            =   180
            TabIndex        =   8
            ToolTipText     =   "Apply semitransparency to sprites"
            Top             =   300
            Width           =   1515
         End
      End
      Begin VB.Frame FramePixel 
         BackColor       =   &H00008000&
         Caption         =   "Pixel Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   4
         Top             =   900
         Width           =   3195
         Begin VB.Frame Frame2 
            BackColor       =   &H0000C000&
            Height          =   795
            Left            =   1500
            TabIndex        =   11
            Top             =   180
            Width           =   1575
            Begin VB.HScrollBar HSFogPix 
               Height          =   255
               Left            =   120
               Max             =   128
               TabIndex        =   15
               Top             =   480
               Value           =   64
               Width           =   915
            End
            Begin VB.CheckBox CheckRedSPR 
               BackColor       =   &H0000C000&
               Caption         =   "R"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   14
               ToolTipText     =   "Select Red component for dotted pixel"
               Top             =   180
               Value           =   1  'Checked
               Width           =   435
            End
            Begin VB.CheckBox CheckGreenSPR 
               BackColor       =   &H0000C000&
               Caption         =   "G"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   600
               TabIndex        =   13
               ToolTipText     =   "Select Green component for dotted pixel"
               Top             =   180
               Value           =   1  'Checked
               Width           =   435
            End
            Begin VB.CheckBox CheckBlueSPR 
               BackColor       =   &H0000C000&
               Caption         =   "B"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   1080
               TabIndex        =   12
               ToolTipText     =   "Select Blue component for dotted pixel"
               Top             =   180
               Value           =   1  'Checked
               Width           =   435
            End
            Begin VB.Label lblFogSPR 
               BackColor       =   &H0000C000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "64"
               Height          =   255
               Left            =   1080
               TabIndex        =   16
               ToolTipText     =   "Degradation value"
               Top             =   480
               Width           =   375
            End
         End
         Begin VB.OptionButton OptPixelXColors 
            BackColor       =   &H00008000&
            Caption         =   "Multi Color"
            Height          =   255
            Left            =   180
            TabIndex        =   6
            ToolTipText     =   "Draw dotted pixels with degradation color"
            Top             =   600
            Width           =   1275
         End
         Begin VB.OptionButton OptPixel1Color 
            BackColor       =   &H00008000&
            Caption         =   "Single Color"
            Height          =   255
            Left            =   180
            TabIndex        =   5
            ToolTipText     =   "Draw dotted pixels with only one color"
            Top             =   300
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.OptionButton OptSprites 
         BackColor       =   &H00008000&
         Caption         =   "Sprites"
         Height          =   315
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Draws dottes with sprites"
         Top             =   540
         Width           =   915
      End
      Begin VB.OptionButton OptPixels 
         BackColor       =   &H00008000&
         Caption         =   "Pixels"
         Height          =   315
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Draws dottes with pixels"
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Unload"
      Height          =   375
      Left            =   8100
      TabIndex        =   0
      Top             =   6600
      Width           =   1035
   End
   Begin VB.Image ImgWIRE 
      Height          =   375
      Index           =   5
      Left            =   9600
      Top             =   2160
      Width           =   2190
   End
   Begin VB.Image ImgWIRE 
      Height          =   375
      Index           =   4
      Left            =   9600
      Top             =   1800
      Width           =   2190
   End
   Begin VB.Image ImgWIRE 
      Height          =   375
      Index           =   3
      Left            =   9600
      Top             =   1440
      Width           =   2190
   End
   Begin VB.Image ImgWIRE 
      Height          =   375
      Index           =   2
      Left            =   9600
      Top             =   1080
      Width           =   2190
   End
   Begin VB.Image ImgWIRE 
      Height          =   375
      Index           =   1
      Left            =   9600
      Top             =   720
      Width           =   2190
   End
   Begin VB.Image imgBOTTOM 
      Height          =   300
      Index           =   1
      Left            =   9600
      Picture         =   "frmMain.frx":0000
      Top             =   2520
      Width           =   2190
   End
   Begin VB.Image ImgWIRE 
      Height          =   375
      Index           =   0
      Left            =   9600
      Top             =   360
      Width           =   2190
   End
   Begin VB.Image imgTOP 
      Height          =   315
      Index           =   1
      Left            =   9600
      Picture         =   "frmMain.frx":05B5
      Top             =   60
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   12
      Left            =   7380
      Top             =   4680
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   11
      Left            =   7380
      Top             =   4320
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   10
      Left            =   7380
      Top             =   3960
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   9
      Left            =   7380
      Top             =   3600
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   8
      Left            =   7380
      Top             =   3240
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   7
      Left            =   7380
      Top             =   2880
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   6
      Left            =   7380
      Top             =   2520
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   5
      Left            =   7380
      Top             =   2160
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   4
      Left            =   7380
      Top             =   1800
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   3
      Left            =   7380
      Top             =   1440
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   2
      Left            =   7380
      Top             =   1080
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   1
      Left            =   7380
      Top             =   720
      Width           =   2190
   End
   Begin VB.Image ImgPIX 
      Height          =   375
      Index           =   0
      Left            =   7380
      Top             =   360
      Width           =   2190
   End
   Begin VB.Image imgBOTTOM 
      Height          =   300
      Index           =   0
      Left            =   7380
      Picture         =   "frmMain.frx":0DE0
      Top             =   5040
      Width           =   2190
   End
   Begin VB.Image imgTOP 
      Height          =   345
      Index           =   0
      Left            =   7380
      Picture         =   "frmMain.frx":1467
      Top             =   60
      Width           =   2190
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":1CE0
      ForeColor       =   &H00FFFFFF&
      Height          =   4275
      Left            =   10140
      TabIndex        =   47
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblFPS 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   7500
      TabIndex        =   45
      ToolTipText     =   "Frames per second"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function RGBsprALL_UNcheck() As Boolean
    If CheckRedSPR.Value = vbUnchecked And _
         CheckGreenSPR.Value = vbUnchecked And _
         CheckBlueSPR.Value = vbUnchecked Then
            If OptPixelXColors.Value = True Then
                OptPixel1Color.Value = True
            End If
            RGBsprALL_UNcheck = True
    Else
        RGBsprALL_UNcheck = False
    End If
End Function
Private Function RGBlinALL_UNcheck() As Boolean
    If CheckRedLIN.Value = vbUnchecked And _
         CheckGreenLIN.Value = vbUnchecked And _
         CheckBlueLIN.Value = vbUnchecked Then
            If OptLineXColors.Value = True Then
                OptLines1Color.Value = True
            End If
            RGBlinALL_UNcheck = True
    Else
        RGBlinALL_UNcheck = False
    End If
End Function


Private Sub CheckAlphaSPR_Click()
    If CheckAlphaSPR.Value = vbChecked Then
        isAlphaEffect = True
        OptIMAX.Enabled = True
    Else
        isAlphaEffect = False
        If OptIMAX.Value = True Then
            OptIMAX.Value = False
            OptSolidSPR.Value = True
        End If
        OptIMAX.Enabled = False
        
    End If
End Sub

Private Sub CheckBackGround_Click()
    If CheckBackGround.Value = vbChecked Then
        isBackGround = True
    Else
        isBackGround = False
    End If
End Sub

Private Sub CheckBlueSPR_Click()
    If RGBsprALL_UNcheck = True Then
        OptPixelXColors.Enabled = False
    Else
        OptPixelXColors.Enabled = True
    End If

    If CheckBlueSPR.Value = vbChecked Then
        isBluePixelDOTT = True
    Else
        isBluePixelDOTT = False
    End If
    If isPixelSingleColorDOTT = False Then
        CreatePalette isRedPixelDOTT, isGreenPixelDOTT, isBluePixelDOTT
    End If
End Sub

Private Sub CheckDiagonals_Click()
    If CheckDiagonals.Value = vbChecked Then
        isNoDiagonals = True
    Else
        isNoDiagonals = False
    End If
    
End Sub

Private Sub CheckGraphFPS_Click()
    If CheckGraphFPS.Value = vbChecked Then
        isGraphFPS = True
        lblFPS.Visible = False
        PicASCHorizontal.Visible = True
    Else
        isGraphFPS = False
        PicASCHorizontal.Visible = False
        lblFPS.Visible = True
    End If
End Sub

Private Sub CheckGreenSPR_Click()
    If RGBsprALL_UNcheck = True Then
        OptPixelXColors.Enabled = False
    Else
        OptPixelXColors.Enabled = True
    End If

    If CheckGreenSPR.Value = vbChecked Then
        isGreenPixelDOTT = True
    Else
        isGreenPixelDOTT = False
    End If
    If isPixelSingleColorDOTT = False Then
        CreatePalette isRedPixelDOTT, isGreenPixelDOTT, isBluePixelDOTT
    End If
End Sub

Private Sub CheckHideFaces_Click()
    If CheckHideFaces.Value = vbChecked Then
        isHidenFacesWIRE = True
    Else
        isHidenFacesWIRE = False
    End If
End Sub

Private Sub CheckQZsortFaces_Click()
    If CheckQZsortFaces.Value = vbChecked Then
        isQZsortWIRE = True
    Else
        isQZsortWIRE = False
    End If
End Sub

Private Sub CheckQZsortPoints_Click()
    If CheckQZsortPoints.Value = vbChecked Then
        isQZsortDOTT = True
    Else
        isQZsortDOTT = False
    End If
End Sub

Private Sub CheckRedSPR_Click()
    If RGBsprALL_UNcheck = True Then
        OptPixelXColors.Enabled = False
    Else
        OptPixelXColors.Enabled = True
    End If

    If CheckRedSPR.Value = vbChecked Then
        isRedPixelDOTT = True
    Else
        isRedPixelDOTT = False
    End If
    If isPixelSingleColorDOTT = False Then
        CreatePalette isRedPixelDOTT, isGreenPixelDOTT, isBluePixelDOTT
    End If
End Sub
Private Sub CheckBlueLIN_Click()
    If RGBlinALL_UNcheck = True Then
        OptLineXColors.Enabled = False
    Else
        OptLineXColors.Enabled = True
    End If
    
    If CheckBlueLIN.Value = vbChecked Then
        isBluePixelWIRE = True
    Else
        isBluePixelWIRE = False
    End If
    If isWireSingleColor = False Then
        CreatePalette isRedPixelWIRE, isGreenPixelWIRE, isBluePixelWIRE
    End If
    
End Sub

Private Sub CheckGreenLIN_Click()
    If RGBlinALL_UNcheck = True Then
        OptLineXColors.Enabled = False
    Else
        OptLineXColors.Enabled = True
    End If
    
    If CheckGreenLIN.Value = vbChecked Then
        isGreenPixelWIRE = True
    Else
        isGreenPixelWIRE = False
    End If
    If isWireSingleColor = False Then
        CreatePalette isRedPixelWIRE, isGreenPixelWIRE, isBluePixelWIRE
    End If
    
End Sub

Private Sub CheckRedLIN_Click()
    If RGBlinALL_UNcheck = True Then
        OptLineXColors.Enabled = False
    Else
        OptLineXColors.Enabled = True
    End If
    
    If CheckRedLIN.Value = vbChecked Then
        isRedPixelWIRE = True
    Else
        isRedPixelWIRE = False
    End If
    If isWireSingleColor = False Then
        CreatePalette isRedPixelWIRE, isGreenPixelWIRE, isBluePixelWIRE
    End If
End Sub

Private Sub CheckShadowedSPR_Click()
    If CheckShadowedSPR.Value = vbChecked Then
        isShadowedSprites = True
    Else
        isShadowedSprites = False
    End If
End Sub

Private Sub CheckSleep_Click()
    If CheckSleep.Value = vbChecked Then
        isSleep = True
    Else
        isSleep = False
    End If
End Sub

Private Sub cmdUnload_Click()
    End
End Sub


Private Sub Form_Load()
Dim i  As Integer
    MsgBox "FASTER if compile this!!!!", vbExclamation, "Info"
    Show
    MakeSinTable
    MakeCosTable
    InitializeAscii
    InitializePlanets
    
    CreatePalette True, True, True

    rHeight = 316
    rWidth = 486

    rHeightASC = 30
    rWidthASC = 130

    XScreen = rWidth 'PicMain.ScaleWidth
    YScreen = rHeight 'PicMain.ScaleHeight
    
    XCenter = (XScreen \ 2) '- 80
    YCenter = YScreen \ 2
    
    isGraphFPS = True
    isSleep = False
    
    isPixelDOTT = True
    isQZsortDOTT = True
    isPixelSingleColorDOTT = True
    isRedPixelDOTT = True
    isGreenPixelDOTT = True
    isBluePixelDOTT = True
    FogValueDOTT = 64
    isAlphaEffect = False
    isSolidSprite = True
    isShadowedSprites = True
    
    isQZsortWIRE = True
    isHidenFacesWIRE = True
    isNoDiagonals = True
    isWireSingleColor = True
    isRedPixelWIRE = True
    isGreenPixelWIRE = True
    isBluePixelWIRE = True
    FogValueWIRE = 64
    
    SpeedXangle = 5
    SpeedYangle = 5
    SpeedZangle = 5
    For i = 0 To 12
        ImgPIX(i).Picture = frmSprites.ImgsOFF(i).Picture
    Next i
    ImgPIX(0).Picture = frmSprites.ImgsON(0).Picture
    WhatPIXrunning = 0
    
    For i = 0 To 5
        ImgWIRE(i).Picture = frmSprites.ImgsWireOFF(i).Picture
    Next i
    WhatWIRErunning = -1
    FirstScene
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RepaintOver
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RepaintOver
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RepaintOver
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RepaintOver
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RepaintOver
End Sub

Private Sub HSFogLin_Change()
    lblFogLIN.Caption = HSFogLin.Value
    FogValueWIRE = HSFogLin.Value
End Sub

Private Sub HSFogLin_Scroll()
    lblFogLIN.Caption = HSFogLin.Value
    FogValueWIRE = HSFogLin.Value
End Sub

Private Sub HSFogPix_Change()
    lblFogSPR.Caption = HSFogPix.Value
    FogValueDOTT = HSFogPix.Value
End Sub

Private Sub HSFogPix_Scroll()
    lblFogSPR.Caption = HSFogPix.Value
    FogValueDOTT = HSFogPix.Value
End Sub

Private Sub HSWidth_Change()
    lblWireWidth.Caption = HSWidth.Value
    frmBack.PicMainWork.DrawWidth = HSWidth.Value
End Sub

Private Sub HSWidth_Scroll()
    lblWireWidth.Caption = HSWidth.Value
    frmBack.PicMainWork.DrawWidth = HSWidth.Value
End Sub

Private Sub HSx_Change()
    lblXaxis.Caption = HSx.Value
    SpeedXangle = HSx.Value
End Sub

Private Sub HSx_Scroll()
    lblXaxis.Caption = HSx.Value
    SpeedXangle = HSx.Value
End Sub

Private Sub HSy_Change()
    lblYaxis.Caption = HSy.Value
    SpeedYangle = HSy.Value
End Sub

Private Sub HSy_Scroll()
    lblYaxis.Caption = HSy.Value
    SpeedYangle = HSy.Value
End Sub

Private Sub HSz_Change()
    lblZaxis.Caption = HSz.Value
    SpeedZangle = HSz.Value
End Sub

Private Sub HSz_Scroll()
    lblZaxis.Caption = HSz.Value
    SpeedZangle = HSz.Value
End Sub



Private Sub ImgPIX_Click(Index As Integer)
Dim i As Integer
    WhatPIXrunning = Index
    WhatWIRErunning = -1
    
    For i = 0 To 5
        ImgWIRE(i).Picture = frmSprites.ImgsWireOFF(i).Picture
    Next i
    
    isBenny = False
    Select Case Index
        Case 0
            ApplyEngine App.Path & "\Meshes\Bola314_1.dVB"
        Case 1
            ApplyEngine App.Path & "\Meshes\Bola314_2.dVB"
        Case 2
            ApplyEngine App.Path & "\Meshes\Torus320_1.dVB"
        Case 3
            ApplyEngine App.Path & "\Meshes\Cube6x6_1.dVB"
        Case 4
            ApplyEngine App.Path & "\Meshes\ADNpdos.dVB"
        Case 5
            ApplyEngine App.Path & "\Meshes\Botellap2.dVB"
        Case 6
            ApplyEngine App.Path & "\Meshes\Cil320p.dVB"
        Case 7
            ApplyEngine App.Path & "\Meshes\Cil320cp.dVB"
        Case 8
            ApplyEngine App.Path & "\Meshes\Copa180.dVB"
        Case 9
            ApplyEngine App.Path & "\Meshes\Misaturnp.dVB"
        Case 10
            ApplyEngine App.Path & "\Meshes\twotorusp.dVB"
        Case 11
            RotateSolarSystem
        Case 12
            isBenny = True
            CheckQZsortPoints.Value = vbUnchecked
            OptSprites.Value = True
            CheckShadowedSPR.Value = vbChecked
            OptSolidSPR.Value = True
            ApplyEngine App.Path & "\Meshes\Benny.dVB"
    End Select

End Sub

Private Sub ImgPIX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RepaintOver
    ImgPIX(Index).Picture = frmSprites.ImgsON(Index).Picture

End Sub

Private Sub ImgWIRE_Click(Index As Integer)
Dim i As Integer
    WhatPIXrunning = -1
    WhatWIRErunning = Index
    For i = 0 To 12
        ImgPIX(i).Picture = frmSprites.ImgsOFF(i).Picture
    Next i
    Select Case Index
        Case 0
            ApplyEngineWire App.Path & "\Meshes\BolaF1.dvb"
        Case 1
            ApplyEngineWire App.Path & "\Meshes\BolaF2.dvb"
        Case 2
            ApplyEngineWire App.Path & "\Meshes\TorusF1.dvb"
        Case 3
            ApplyEngineWire App.Path & "\Meshes\cubewire.dvb"
        Case 4
            ApplyEngineWire App.Path & "\Meshes\ADNpdos.dvb"
        Case 5
            ApplyEngineWire App.Path & "\Meshes\copa180dos.dvb"
    End Select
End Sub

Private Sub ImgWIRE_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RepaintOver
    ImgWIRE(Index).Picture = frmSprites.ImgsWireON(Index).Picture
End Sub

Private Sub OptIMAX_Click()
    isSolidSprite = False
End Sub

Private Sub OptLines1Color_Click()
    isWireSingleColor = True
End Sub

Private Sub OptLineXColors_Click()
    isWireSingleColor = False
    CreatePalette isRedPixelWIRE, isGreenPixelWIRE, isBluePixelWIRE
End Sub

Private Sub OptPixel1Color_Click()
    isPixelSingleColorDOTT = True
End Sub

Private Sub OptPixels_Click()
    FrameSpr.Enabled = False
    FramePixel.Enabled = True
    isPixelDOTT = True
End Sub

Private Sub OptPixelXColors_Click()
    isPixelSingleColorDOTT = False
    CreatePalette isRedPixelDOTT, isGreenPixelDOTT, isBluePixelDOTT
End Sub

Private Sub OptSolidSPR_Click()
    isSolidSprite = True
End Sub

Private Sub OptSprites_Click()
    FrameSpr.Enabled = True
    FramePixel.Enabled = False
    isPixelDOTT = False
End Sub



Private Sub PicMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RepaintOver
End Sub
Private Sub RepaintOver()
Dim i As Integer
    For i = 0 To 12
        ImgPIX(i).Picture = frmSprites.ImgsOFF(i).Picture
    Next i
    If WhatPIXrunning <> -1 Then
        ImgPIX(WhatPIXrunning).Picture = frmSprites.ImgsON(WhatPIXrunning).Picture
    End If

    For i = 0 To 5
        ImgWIRE(i).Picture = frmSprites.ImgsWireOFF(i).Picture
    Next i
    If WhatWIRErunning <> -1 Then
        ImgWIRE(WhatWIRErunning).Picture = frmSprites.ImgsWireON(WhatWIRErunning).Picture
    End If
End Sub
