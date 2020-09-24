VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rotate Settings"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   690
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame frameParent 
      Caption         =   "Flip or rotate"
      Height          =   2925
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3165
      Begin VB.Frame frameChild 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1275
         Left            =   360
         TabIndex        =   6
         Top             =   1560
         Width           =   2595
         Begin VB.OptionButton optDegree 
            Caption         =   "&270 Degree"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   570
            TabIndex        =   9
            Top             =   810
            Width           =   1500
         End
         Begin VB.OptionButton optDegree 
            Caption         =   "&180 Degree"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   570
            TabIndex        =   8
            Top             =   435
            Width           =   1500
         End
         Begin VB.OptionButton optDegree 
            Caption         =   "&90 Degree"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   570
            TabIndex        =   7
            Top             =   60
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "&Rotate by angle"
         Height          =   285
         Index           =   2
         Left            =   330
         TabIndex        =   5
         Top             =   1200
         Width           =   1500
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "Flip &vertical"
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   4
         Top             =   810
         Width           =   1500
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "&Flip horizontal"
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   420
         Value           =   -1  'True
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function ChangeSelection(nVal As Integer)
    Select Case nVal
        Case 0
            optDegree(0).Enabled = False
            optDegree(1).Enabled = False
            optDegree(2).Enabled = False
            
        Case 1
            optDegree(0).Enabled = False
            optDegree(1).Enabled = False
            optDegree(2).Enabled = False
        
        Case 2
            optDegree(0).Enabled = True
            optDegree(1).Enabled = True
            optDegree(2).Enabled = True
    End Select
End Function

Private Sub btnCancel_Click()
    blnFlag = False
    Unload Me
End Sub

Private Sub btnOK_Click()
    blnFlag = True
    Unload Me
End Sub

Private Sub optFlip_Click(Index As Integer)
    ChangeSelection Index
End Sub
