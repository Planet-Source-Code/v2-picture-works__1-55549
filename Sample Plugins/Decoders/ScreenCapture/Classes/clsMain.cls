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
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function GoGetIt(PicDC As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional Reserved As Variant) As Boolean
Dim sTmp As String
        
    Load frmSettings
    frmSettings.Show
    Do
        DoEvents
    Loop Until frmSettings.Visible = False
    
    If DoneIt = False Then Exit Function
    If WhenToCapture = 0 Then Sleep (3000)
    sTmp = App.Path & "\" & "ScreenCapture_" & Format(Now, "HHMMSSDDMMYYYY") & ".bmp"
    
    If WhatToCapture = 0 Then
        Call SavePicture(CaptureActiveWindow, sTmp)
    Else
        Call SavePicture(CaptureScreen, sTmp)
    End If
    
    Reserved = sTmp
    
End Function
