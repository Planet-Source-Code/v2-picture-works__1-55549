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
    Dim XORColor As Long
    
    w = X2 - X1
    h = Y2 - Y1
    
    For x = 0 To w
    DoEvents
        For y = 0 To h
            lngColor = GetPixel(PicDC, x, y)
            lngColor = lngColor Xor RGB(255, 0, 0)
            lngColor = lngColor Xor RGB(0, 255, 0)
            lngColor = lngColor Xor RGB(0, 0, 255)
            SetPixel PicDC, x, y, lngColor
        Next y
    Next x
    GoGetIt=True        
End Function
