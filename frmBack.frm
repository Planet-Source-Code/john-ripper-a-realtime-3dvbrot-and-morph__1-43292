VERSION 5.00
Begin VB.Form frmBack 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   367
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   708
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicASCHorizontalWork 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   8580
      ScaleHeight     =   30
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   211.25
      TabIndex        =   3
      Top             =   420
      Width           =   1950
   End
   Begin VB.PictureBox PicASCHorizontalClean 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   8160
      ScaleHeight     =   30
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   211.25
      TabIndex        =   2
      Top             =   180
      Width           =   1950
   End
   Begin VB.PictureBox PicMainWork 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   4740
      Left            =   720
      ScaleHeight     =   316
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   486
      TabIndex        =   1
      Top             =   480
      Width           =   7290
   End
   Begin VB.PictureBox PicMainClean 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4740
      Left            =   120
      ScaleHeight     =   316
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   486
      TabIndex        =   0
      Top             =   0
      Width           =   7290
   End
End
Attribute VB_Name = "frmBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

