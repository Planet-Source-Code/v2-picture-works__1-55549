VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Function GoGetIt(PicDC As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional Reserved As Variant) As Boolean
    Dim w As Long
    Dim h As Long
    Dim p(2) As POINTAPI
    Dim zAng As Long
    Dim zCos As Long
    Dim zSin As Long
    Dim ixc1 As Long
    Dim ixc2 As Long
    Dim iyc1 As Long
    Dim iyc2 As Long
    Dim ix As Long
    Dim iy As Long
    Dim xs As Long
    Dim xy As Long
    Dim ixs0 As Long
    Dim iys0 As Long
    Dim xfs1 As Long
    Dim yfs1 As Long
    Dim longcul0 As Long
    Dim longcul1 As Long
    Dim tmpPic As StdPicture
    Set tmpPic = New StdPicture
    
    Const pi# = 3.1415926

    w = X2 - X1
    h = Y2 - Y1
    

'BitBlt PicDC, 0, 0, w, h, PicDC, 0, 0, vbSrcErase
zAng = pi# / 6 ' Rotation angle

zCos = Cos(zAng)
zSin = Sin(zAng)

PicW1 = w
PicH1 = h
picW2 = w
picH2 = h

ixc1 = w / 2: iyc1 = h / 2
ixc2 = w / 2: iyc2 = h / 2

For iy = 0 To picH2 - 1
For ix = 0 To picW2 - 1

   xs = ixc1 + (ix - ixc2) * zCos + (iy - iyc2) * zSin
   ys = iyc1 + (iy - iyc2) * zCos - (ix - ixc2) * zSin

   ' Bottom left coords of bounding
   ' rectangle on Picture1
   ixs0 = Int(xs)
   iys0 = Int(ys)
   
   If ixs0 > 0 And ixs0 < PicW1 - 1 And iys0 > 0 And iys0 < PicH1 - 1 Then
      
      xfs1 = xs - Int(xs)
      yfs1 = ys - Int(ys)
      
      'ixs0->ixs0+1, iyso
      longcul0 = GetPixel(PicDC, ixs0, iys0)
      redc = GetRed(longcul0)
      greenc = GetGreen(longcul0)
      bluec = GetBlue(longcul0)
      
      culr = (1 - xfs1) * redc
      culg = (1 - xfs1) * greenc
      culb = (1 - xfs1) * bluec
   
      longcul1 = GetPixel(PicDC, ixs0 + 1, iys0)
      redc = GetRed(longcul1)
      greenc = GetGreen(longcul1)
      bluec = GetBlue(longcul1)
      
      culr0 = culr + xfs1 * redc
      culg0 = culg + xfs1 * greenc
      culb0 = culb + xfs1 * bluec
      
      'ixs0->ixs0+1, iys0+1
      longcul0 = GetPixel(PicDC, ixs0, iys0 + 1)
      redc = GetRed(longcul0)
      greenc = GetGreen(longcul0)
      bluec = GetBlue(longcul0)
      
      culr = (1 - xfs1) * redc
      culg = (1 - xfs1) * greenc
      culb = (1 - xfs1) * bluec
      
      longcul1 = GetPixel(PicDC, ixs0 + 1, iys0 + 1)
      redc = GetRed(longcul1)
      greenc = GetGreen(longcul1)
      bluec = GetBlue(longcul1)
      
      culr1 = culr + xfs1 * redc
      culg1 = culg + xfs1 * greenc
      culb1 = culb + xfs1 * bluec
   
      ' Weight along y axis
      culr = (1 - yfs1) * culr0 + yfs1 * culr1
      culg = (1 - yfs1) * culg0 + yfs1 * culg1
      culb = (1 - yfs1) * culb0 + yfs1 * culb1
   
      res& = SetPixel(PicDC, ix, iy, RGB(culr, culg, culb))
      
   End If
   
Next ix
Next iy
    
    GoGetIt = True
End Function

Public Function GetRed(lngColor As Long) As Long
On Error Resume Next
    
    Dim bRed As Byte
    
    bRed = lngColor And &HFF
    If bRed < 0 Then bRed = 0
    If bRed > 255 Then bRed = 255
    GetRed = bRed

End Function

Public Function GetGreen(lngColor As Long) As Long
On Error Resume Next
    
    Dim bGreen As Byte
    
    bGreen = ((lngColor And &HFF00) / &H100) Mod &H100
    If bGreen < 0 Then bGreen = 0
    If bGreen > 255 Then bGreen = 255
    GetGreen = bGreen

End Function
Public Function GetBlue(lngColor As Long) As Long
    
On Error Resume Next
    
    Dim bBlue As Byte
    
    bBlue = ((lngColor And &HFF0000) / &H10000) Mod &H100
    If bBlue < 0 Then bBlue = 0
    If bBlue > 255 Then bBlue = 255
    GetBlue = bBlue

End Function




