VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Darken Settings"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   285
      Left            =   2790
      TabIndex        =   3
      Top             =   690
      Width           =   885
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Min             =   1
      Max             =   255
      SelStart        =   50
      TickFrequency   =   10
      Value           =   50
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   30
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   210
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOk_Click()
    Unload Me
End Sub


