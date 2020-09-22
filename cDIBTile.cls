VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// -------------------
'// Class : cDIBTile
'// Author: Carles P.V.
'// -------------------
'// Date Cr: 10.11.2002
'// Last Md: 08.02.2003
'// -------------------


Option Explicit

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const OBJ_BITMAP     As Long = 7

Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'//

Private m_hBrush As Long '* Pattern brush

'//

Private Sub Class_Initialize()
    m_hBrush = 0
End Sub

Public Function SetPattern(Picture As StdPicture) As Boolean

  Dim tBI       As BITMAP
  Dim tBIH      As BITMAPINFOHEADER
  Dim Buff()    As Byte 'Packed DIB
    
  Dim lhDC      As Long
  Dim lhOldBmp  As Long
    
    If (GetObjectType(Picture) = OBJ_BITMAP) Then
    
        '-- Get image info
        GetObject Picture, Len(tBI), tBI
        
        '-- Prepare DIB header and redim. Buff() array
        With tBIH
            .biSize = Len(tBIH) '40
            .biPlanes = 1
            .biBitCount = 24
            .biWidth = tBI.bmWidth
            .biHeight = tBI.bmHeight
            .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        End With
        ReDim Buff(1 To Len(tBIH) + tBIH.biSizeImage) '[Header + Bits]
            
        '-- Create DIB brush
        lhDC = CreateCompatibleDC(0)
        If (lhDC <> 0) Then
            lhOldBmp = SelectObject(lhDC, Picture)
                    
            '-- Build packed DIB:
            '   - Merge Header
                CopyMemory Buff(1), tBIH, Len(tBIH)
            '   - Get and merge DIB Bits
                GetDIBits lhDC, Picture, 0, tBI.bmHeight, Buff(Len(tBIH) + 1), tBIH, DIB_RGB_COLORS
            
            SelectObject lhDC, lhOldBmp
            DeleteDC lhDC
            
            '-- Create brush from packed DIB
            DestroyPattern
            m_hBrush = CreateDIBPatternBrushPt(Buff(1), DIB_RGB_COLORS)
        End If
      Else
        Debug.Print "> Picture is empty"
    End If
    
    SetPattern = (m_hBrush <> 0)
End Function

Public Sub Tile(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

  Dim TileRect As RECT2
  Dim PtOrg    As POINTAPI
  
    If (m_hBrush <> 0) Then
        SetRect TileRect, x1, y1, x2, y2
        SetBrushOrgEx hdc, x1, y1, PtOrg
        '-- Tile image
        FillRect hdc, TileRect, m_hBrush
      Else
        Debug.Print "> Pattern brush has not been initialized"
    End If
End Sub

Public Sub DestroyPattern()
    If (m_hBrush <> 0) Then
        DeleteObject m_hBrush
        m_hBrush = 0
    End If
End Sub

'//

Private Sub Class_Terminate()
    DestroyPattern
End Sub
