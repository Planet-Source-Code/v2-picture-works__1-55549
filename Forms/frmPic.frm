VERSION 5.00
Begin VB.Form frmPic 
   Caption         =   "V2 Softwares"
   ClientHeight    =   3600
   ClientLeft      =   3285
   ClientTop       =   2850
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   Begin VB.HScrollBar hScroll 
      Height          =   195
      LargeChange     =   5
      Left            =   150
      Max             =   100
      Min             =   1
      TabIndex        =   3
      Top             =   2310
      Value           =   1
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.VScrollBar vScroll 
      Height          =   1245
      LargeChange     =   5
      Left            =   4800
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   60
      Value           =   1
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   150
      Picture         =   "frmPic.frx":0000
      ScaleHeight     =   148
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   307
      TabIndex        =   1
      Top             =   60
      Width           =   4635
   End
   Begin VB.FileListBox PlugFile 
      Height          =   480
      Left            =   3360
      TabIndex        =   0
      Top             =   2700
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LeftExtra As Long
Dim TopExtra As Long


Public Sub Form_Resize()
    
    If WindowState = vbMinimized Then Exit Sub
    If ScaleWidth < 100 Or ScaleHeight < 100 Then Exit Sub
    
    If picMain.Width > ScaleWidth Then
        hScroll.Move 3, ScaleHeight - hScroll.Height - 3, ScaleWidth - 6
        hScroll.Visible = True
        GetLeft
    Else
        hScroll.Visible = False
        picMain.Left = (ScaleWidth / 2) - (picMain.ScaleWidth / 2)
    End If
        
    If picMain.Height > ScaleHeight Then
        vScroll.Move ScaleWidth - vScroll.Width - 6, 3, vScroll.Width, ScaleHeight - 6
        vScroll.Visible = True
        picMain.Top = GetTop
    Else
        vScroll.Visible = False
        picMain.Top = (ScaleHeight / 2) - (picMain.ScaleHeight / 2)
    End If
End Sub

Private Function GetLeft() As Long

    LeftExtra = (picMain.Width - ScaleWidth) * -1
    hScroll.Max = LeftExtra - 30
End Function

Private Function GetTop() As Long
    TopExtra = (picMain.Height - ScaleHeight) * -1
    vScroll.Max = TopExtra - 30
End Function

Private Sub hScroll_Change()
    hScroll_Scroll
End Sub

Private Sub hScroll_Scroll()
    picMain.Left = (LeftExtra + hScroll.Value) + 30
End Sub

Private Sub vScroll_Change()
    vScroll_Scroll
End Sub

Private Sub vScroll_Scroll()
    picMain.Top = (TopExtra + vScroll.Value) + 30
End Sub
