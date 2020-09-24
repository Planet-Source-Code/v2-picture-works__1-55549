VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgVersion 
      Height          =   420
      Left            =   1470
      Picture         =   "frmSplash.frx":000C
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":1B6E
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1425
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   1860
      Width           =   5055
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.0.1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   1290
      Width           =   735
   End
   Begin VB.Image imgName 
      Height          =   1215
      Left            =   270
      Picture         =   "frmSplash.frx":1C62
      Top             =   90
      Width           =   4740
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2070
      TabIndex        =   0
      Top             =   3450
      Width           =   1200
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.0.1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   495
      Index           =   1
      Left            =   2790
      TabIndex        =   2
      Top             =   1305
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   -30
      Top             =   3390
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":14898
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   1425
      Index           =   1
      Left            =   190
      TabIndex        =   4
      Top             =   1870
      Width           =   5055
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Click()
    Hide
End Sub


Private Sub Form_Load()
    lblVersion(0) = App.Major & "." & App.Minor & "." & App.Revision
    lblVersion(1) = lblVersion(0)
    Show
    StayOnTop Me, True
    EnumPlugins
    StayOnTop Me, False
    Hide
    frmMain.Show
    frmMain.Enabled = True
End Sub

Private Sub imgName_Click()
    Hide
End Sub

Private Sub imgVersion_Click()
    Hide
End Sub

Private Sub Label1_Click(index As Integer)
    Hide
End Sub

Private Sub lblStatus_Click()
    Hide
End Sub

Private Sub lblVersion_Click(index As Integer)
    Hide
End Sub
