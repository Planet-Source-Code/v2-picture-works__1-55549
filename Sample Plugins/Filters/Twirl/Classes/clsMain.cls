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

Function GoGetIt(PicDC As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional Reserved As Variant) As Boolean
    Dim x As Long
    Dim y As Long
    Dim w As Long
    Dim h As Long
    Dim lngColor As Long
    Dim lngColor1 As Long
    Dim R As Integer
    Dim g As Integer
    Dim B As Integer
    Dim lngVal As Byte
    
    w = X2 - X1
    h = Y2 - Y1
    
    frmSettings.Show
    
    Do
        DoEvents
    Loop While frmSettings.Visible = True
    If blnOK = False Then Exit Function
    Render_Twirl PicDC, w, h, nVal - 100
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



