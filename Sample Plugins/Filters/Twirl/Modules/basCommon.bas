Attribute VB_Name = "basCommon"
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Global blnOK As Boolean
Global nVal As Integer

Public Function Render_Twirl(DC As Long, w As Long, h As Long, Angle As Double)
    Dim Rad As Double
    Dim A As Double
    Dim B As Double
    Dim x As Double
    Dim y As Double
    Dim R As Double
    Dim C As Long

    Dim OS_Y As Integer
    
    Const PI = 3.1415
    
    x = w / 2
    y = h / 2
    
    B = Angle / ((w / 2) * 100)
    
    For Rad = y To 0 Step -0.1
        A = A + B
        For R = 0 To PI * 2 Step (w / 2) / ((w / 2) * 100)
            C = GetPixel(DC, (x + Cos(R) * Rad), (y + Sin(R) * Rad))
            SetPixel DC, (x + Cos(A + R) * Rad), (y + Sin(A + R) * Rad), C
        Next R
        DoEvents
    Next Rad

End Function


