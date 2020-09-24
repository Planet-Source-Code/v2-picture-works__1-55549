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

Function GoGetIt(PicDC As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional Reserved As Variant) As Boolean
    Dim x As Long
    Dim y As Long
    Dim w As Long
    Dim h As Long
    Dim lngColor As Long
    Dim lngColor1 As Long
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    
    
    w = X2 - X1
    h = Y2 - Y1
    
    For x = 0 To w
    DoEvents
        For y = 0 To h
            lngColor = GetPixel(PicDC, x, y)
            lngColor1 = GetPixel(PicDC, x + 1, y)
            r = Abs((lngColor1 And &HFF) + (lngColor And &HFF)) / 2
            g = Abs(((lngColor1 And &HFF00) / &H100) Mod &H100 + ((lngColor And &HFF00) / &H100) Mod &H100) / 2
            b = Abs(((lngColor1 And &HFF0000) / &H10000) Mod &H100 + ((lngColor And &HFF0000) / &H10000) Mod &H100) / 2
            SetPixel PicDC, x, y, RGB(r, g, b)
            
        Next y
    Next x
    GoGetIt=True
End Function
