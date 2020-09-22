VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "DIB Brush test"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private cTile As New cDIBTile

Private Sub Form_Load()
    
    '-- Set pattern
    cTile.SetPattern LoadResPicture("PSC", vbResBitmap)
End Sub

Private Sub Form_Paint()
    
    '-- Tile pattern
    cTile.Tile hdc, 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set cTile = Nothing
    Set fTest = Nothing
End Sub
