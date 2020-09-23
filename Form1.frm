VERSION 5.00
Object = "*\ALCDisplay.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin LCDisplay.Display Display1 
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   2160
      _ExtentX        =   19050
      _ExtentY        =   582
      Characters      =   9
      Caption         =   "Display"
      ScrollRate      =   1
      Scroll          =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------
' Fernando Zamora
'
' LED Marquee/Label (Tuned)
' This code is only a high performance of the original made by Wasp53x
' i only change some functions so all the work is from the autor, i just though
' is was a great code but a bit slowly so dont wait any more and rate the original autor
' at
' www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=25019
