VERSION 5.00
Begin VB.Form frmSprites 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   598
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   1016
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSprBenny 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   8160
      Picture         =   "frmSprites.frx":0000
      ScaleHeight     =   28
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   68
      TabIndex        =   4
      Top             =   2760
      Width           =   1020
   End
   Begin VB.PictureBox PicSPRascii 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Left            =   -6660
      Picture         =   "frmSprites.frx":0C42
      ScaleHeight     =   173
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   1009
      TabIndex        =   3
      Top             =   3060
      Width           =   15195
   End
   Begin VB.PictureBox PicMainBackGround 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Left            =   1740
      Picture         =   "frmSprites.frx":80D7C
      ScaleHeight     =   311.078
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   486
      TabIndex        =   2
      Top             =   480
      Width           =   7290
   End
   Begin VB.PictureBox PicSPR3d 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7320
      Picture         =   "frmSprites.frx":8465A
      ScaleHeight     =   25
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   53
      TabIndex        =   1
      Top             =   2760
      Width           =   795
   End
   Begin VB.PictureBox PicPlanets 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   5160
      Picture         =   "frmSprites.frx":84E1C
      ScaleHeight     =   133
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   273
      TabIndex        =   0
      Top             =   6060
      Width           =   4095
   End
   Begin VB.Image ImgsWireOFF 
      Height          =   375
      Index           =   5
      Left            =   13020
      Picture         =   "frmSprites.frx":9BE4E
      Top             =   3600
      Width           =   2190
   End
   Begin VB.Image ImgsWireOFF 
      Height          =   375
      Index           =   4
      Left            =   13020
      Picture         =   "frmSprites.frx":9CC6C
      Top             =   3240
      Width           =   2190
   End
   Begin VB.Image ImgsWireOFF 
      Height          =   375
      Index           =   3
      Left            =   13020
      Picture         =   "frmSprites.frx":9DA8F
      Top             =   2880
      Width           =   2190
   End
   Begin VB.Image ImgsWireON 
      Height          =   375
      Index           =   5
      Left            =   12060
      Picture         =   "frmSprites.frx":9E8D4
      Top             =   4200
      Width           =   2190
   End
   Begin VB.Image ImgsWireON 
      Height          =   375
      Index           =   4
      Left            =   12060
      Picture         =   "frmSprites.frx":9F6C3
      Top             =   3840
      Width           =   2190
   End
   Begin VB.Image ImgsWireON 
      Height          =   375
      Index           =   3
      Left            =   12060
      Picture         =   "frmSprites.frx":A04BE
      Top             =   3480
      Width           =   2190
   End
   Begin VB.Image ImgsWireOFF 
      Height          =   375
      Index           =   2
      Left            =   13020
      Picture         =   "frmSprites.frx":A12C0
      Top             =   2520
      Width           =   2190
   End
   Begin VB.Image ImgsWireOFF 
      Height          =   375
      Index           =   1
      Left            =   13020
      Picture         =   "frmSprites.frx":A212B
      Top             =   2160
      Width           =   2190
   End
   Begin VB.Image ImgsWireOFF 
      Height          =   375
      Index           =   0
      Left            =   13020
      Picture         =   "frmSprites.frx":A307B
      Top             =   1800
      Width           =   2190
   End
   Begin VB.Image ImgsWireON 
      Height          =   375
      Index           =   2
      Left            =   12060
      Picture         =   "frmSprites.frx":A3EF3
      Top             =   3120
      Width           =   2190
   End
   Begin VB.Image ImgsWireON 
      Height          =   375
      Index           =   1
      Left            =   12060
      Picture         =   "frmSprites.frx":A4D3E
      Top             =   2760
      Width           =   2190
   End
   Begin VB.Image ImgsWireON 
      Height          =   375
      Index           =   0
      Left            =   12060
      Picture         =   "frmSprites.frx":A5C40
      Top             =   2400
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   3
      Left            =   10260
      Picture         =   "frmSprites.frx":A6A90
      Top             =   1500
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   12
      Left            =   10260
      Picture         =   "frmSprites.frx":A78F8
      Top             =   4740
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   11
      Left            =   10260
      Picture         =   "frmSprites.frx":A8769
      Top             =   4380
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   10
      Left            =   10260
      Picture         =   "frmSprites.frx":A9741
      Top             =   4020
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   9
      Left            =   10260
      Picture         =   "frmSprites.frx":AA630
      Top             =   3660
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   8
      Left            =   10260
      Picture         =   "frmSprites.frx":AB490
      Top             =   3300
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   7
      Left            =   10260
      Picture         =   "frmSprites.frx":AC2A0
      Top             =   2940
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   6
      Left            =   10260
      Picture         =   "frmSprites.frx":AD248
      Top             =   2580
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   5
      Left            =   10260
      Picture         =   "frmSprites.frx":AE1F2
      Top             =   2220
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   4
      Left            =   10260
      Picture         =   "frmSprites.frx":AF05B
      Top             =   1860
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   16
      Left            =   10260
      Picture         =   "frmSprites.frx":AFE96
      Top             =   1500
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   2
      Left            =   10260
      Picture         =   "frmSprites.frx":B0CFE
      Top             =   1140
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   1
      Left            =   10260
      Picture         =   "frmSprites.frx":B1B70
      Top             =   780
      Width           =   2190
   End
   Begin VB.Image ImgsOFF 
      Height          =   375
      Index           =   0
      Left            =   10260
      Picture         =   "frmSprites.frx":B2B6F
      Top             =   420
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   12
      Left            =   9540
      Picture         =   "frmSprites.frx":B3A32
      Top             =   7140
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   11
      Left            =   9540
      Picture         =   "frmSprites.frx":B489D
      Top             =   6780
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   10
      Left            =   9540
      Picture         =   "frmSprites.frx":B580C
      Top             =   6420
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   9
      Left            =   9540
      Picture         =   "frmSprites.frx":B66ED
      Top             =   6060
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   8
      Left            =   9540
      Picture         =   "frmSprites.frx":B7561
      Top             =   5700
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   7
      Left            =   9540
      Picture         =   "frmSprites.frx":B838E
      Top             =   5340
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   6
      Left            =   9540
      Picture         =   "frmSprites.frx":B92F3
      Top             =   4980
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   5
      Left            =   9540
      Picture         =   "frmSprites.frx":BA261
      Top             =   4620
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   4
      Left            =   9540
      Picture         =   "frmSprites.frx":BB0C1
      Top             =   4260
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   3
      Left            =   9540
      Picture         =   "frmSprites.frx":BBEF9
      Top             =   3900
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   2
      Left            =   9540
      Picture         =   "frmSprites.frx":BCD70
      Top             =   3540
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   1
      Left            =   9540
      Picture         =   "frmSprites.frx":BDBF7
      Top             =   3180
      Width           =   2190
   End
   Begin VB.Image ImgsON 
      Height          =   375
      Index           =   0
      Left            =   9540
      Picture         =   "frmSprites.frx":BEB66
      Top             =   2820
      Width           =   2190
   End
End
Attribute VB_Name = "frmSprites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

