VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Twirl Settings"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnApply 
      Caption         =   "Apply"
      Height          =   255
      Left            =   1500
      TabIndex        =   4
      Top             =   2640
      Width           =   795
   End
   Begin VB.PictureBox TwirlPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   2235
      Left            =   60
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   60
      Width           =   2235
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   60
      Max             =   200
      TabIndex        =   1
      Top             =   2340
      Value           =   100
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Twirl Value:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   2700
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Private Sub btnApply_Click()
    blnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Draw_Twirl 0
End Sub


Private Sub HScroll1_Change()
    Draw_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.Label1.Caption = Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.TwirlPic.SetFocus
    nVal = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    Draw_Twirl Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.Label1.Caption = Me.HScroll1.Value - (Me.HScroll1.Max / 2)
    Me.TwirlPic.SetFocus
    nVal = HScroll1.Value
End Sub


'DRAW TWIRL
Public Sub Draw_Twirl(Angle As Double)
    Dim Rad As Double
    Dim A As Double
    Dim B As Double
    Dim x As Double
    Dim y As Double
    
    x = Me.TwirlPic.ScaleWidth / 2
    y = Me.TwirlPic.ScaleHeight / 2
    
    B = Angle / 10000
    
    Me.TwirlPic.Cls
    
    For Rad = 100 To 0 Step -0.1
        A = A + B
        SetPixelV Me.TwirlPic.hdc, x + Cos(A) * Rad, y + Sin(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x - Cos(A) * Rad, y - Sin(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x - Sin(A) * Rad, y + Cos(A) * Rad, 0
        SetPixelV Me.TwirlPic.hdc, x + Sin(A) * Rad, y - Cos(A) * Rad, 0
    Next Rad
    
    
    
End Sub
