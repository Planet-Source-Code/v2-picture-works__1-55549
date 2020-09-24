VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Picture Works"
   ClientHeight    =   4995
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6915
   Enabled         =   0   'False
   Icon            =   "frmMain.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   4755
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9128
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:08 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   3060
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "Close A&ll"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "Filte&rs"
      Begin VB.Menu mnuPlug 
         Caption         =   "No Filters Found"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPlugManager 
         Caption         =   "Plugin &Manager"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowTileH 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWindowTileV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowsSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "Open Windows Here"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutPlugins 
         Caption         =   "About Plugins"
         Begin VB.Menu mnuAboutPlug 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuAboutSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuAboutApp 
         Caption         =   "Picture Works"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_DblClick()
    mnuFileOpen_Click
End Sub

Private Sub MDIForm_Load()
    Me.Visible = False
    Caption = "Picture Works - [ Version " & App.Major & "." & App.Minor & "." & App.Revision & " ]"
    Status.Panels(1).Text = "Ready"
    Unload frmPic
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuAboutApp_Click()
    frmSplash.Show
End Sub

Private Sub mnuAboutPlug_Click(index As Integer)
    Call Plugs(index - 1).objInfo.About
End Sub

Private Sub mnuFileClose_Click()
On Error Resume Next
    Set ActiveForm.picMain = Nothing
    Unload ActiveForm
    nTotalForms = nTotalForms - 1
End Sub

Private Sub mnuFileCloseAll_Click()
On Error Resume Next
    For Each Form In Me
        Unload ActiveForm
        Status.Panels(1).Text = "Closing " & ActiveForm.Caption
        Status.Panels(1).Text = "Ready"
    Next
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileOpen_Click()
On Error Resume Next
Dim n As Integer
Dim nDecoder As Integer

Dim sFiles() As String
    With cmd
        .Flags = CMD_ALLOWMULTISELECT Or CMD_EXPLORER
        .DialogTitle = n
        .Filter = GetDecoderList
        .FileName = ""
        .ShowOpen
        sFileName = .FileName
    End With
    
    Select Case cmd.FilterIndex - 1
    Case 0
        
        If Len(sFileName) <= 0 Then Exit Sub
        sFiles = Split(sFileName, Chr(0))
        If UBound(sFiles) > 0 Then
            For n = 1 To UBound(sFiles)
                ReDim Preserve objForm(nTotalForms) As Form
                Set objForm(nTotalForms) = New frmPic
                Load objForm(nTotalForms)
                objForm(nTotalForms).picMain = LoadPicture(sFiles(0) & "\" & sFiles(n))
                objForm(nTotalForms).Caption = sFiles(n)
                objForm(nTotalForms).Form_Resize
                nTotalForms = nTotalForms + 1
            Next n
        ElseIf UBound(sFiles) = 0 Then
                ReDim Preserve objForm(nTotalForms) As Form
                Set objForm(nTotalForms) = New frmPic
                Load objForm(nTotalForms)
                objForm(nTotalForms).picMain = LoadPicture(sFiles(0))
                objForm(nTotalForms).Caption = Mid(sFiles(0), InStrRev(sFiles(0), "\") + 1, Len(sFiles(0)) - InStrRev(sFiles(0), "\") + 1)
                objForm(nTotalForms).Form_Resize
                nTotalForms = nTotalForms + 1
        End If
Case Else
    ReDim Preserve objForm(nTotalForms) As Form
    Set objForm(nTotalForms) = New frmPic
    Load objForm(nTotalForms)
    
    nDecoder = GetDecoder(cmd.FilterIndex)
    
    Call Plugs(nDecoder).objMain.GoGetIt(0, 0, 0, 0, 0, sFileName)
    objForm(nTotalForms).picMain = LoadPicture(sFileName)
    Kill sFileName
    objForm(nTotalForms).Caption = "Picture Loaded By " & Plugs(nDecoder).objDetails.Name
    objForm(nTotalForms).Form_Resize
    nTotalForms = nTotalForms + 1
End Select
        
End Sub

Private Sub mnuFileSave_Click()
Dim nPlug As Integer
    If Len(sFileName) <= 0 Then
        cmd.Filter = GetEncoderList
        cmd.ShowSave
        sFileName = cmd.FileName
    End If
    Select Case cmd.FilterIndex - 1
        Case 0
            SavePicture frmPic.picMain.Image, sFileName
        Case Else
            nPlug = GetEncoder(cmd.FilterIndex)
            Call Plugs(nPlug).objMain.GoGetIt(ActiveForm.picMain.hDC, 0, 0, ActiveForm.picMain.ScaleWidth, ActiveForm.picMain.ScaleHeight, sFileName)
    End Select
        
End Sub

Private Sub mnuFileSaveAs_Click()
    On Error Resume Next
    Dim nPlug As Integer
    
    cmd.Filter = GetEncoderList
    cmd.ShowSave
    sFileName = cmd.FileName
    
    Select Case cmd.FilterIndex - 1
        Case 0
            SavePicture frmPic.picMain.Image, sFileName
        Case Else
            nPlug = GetEncoder(cmd.FilterIndex)
            Call Plugs(nPlug).objMain.GoGetIt(ActiveForm.picMain.hDC, 0, 0, ActiveForm.picMain.ScaleWidth, ActiveForm.picMain.ScaleHeight, sFileName)
    End Select
End Sub

Private Sub mnuPlug_Click(index As Integer)
On Error Resume Next
    Dim lngReturn As Long
    Dim objPic As StdPicture
    Dim nFilter As Integer
    
    Set objPic = ActiveForm.picMain.Picture
    
    nFilter = GetFilter(Trim(Right(mnuPlug(index).Caption, Len(mnuPlug(index).Caption) - 4)))
    
    Status.Panels(1).Text = "Applying Plugin " & Plugs(nFilter).Name
    lngReturn = Plugs(nFilter).objMain.GoGetIt(ActiveForm.picMain.hDC, 0, 0, ActiveForm.picMain.ScaleWidth, ActiveForm.picMain.ScaleHeight, Me.hWnd)
    If lngReturn = False Then ActiveForm.picMain.Picture = objPic
    ActiveForm.picMain.Refresh
    Status.Panels(1).Text = "Ready"
    Set objPic = Nothing
End Sub

Private Sub mnuPlugManager_Click()
    frmPlugManager.Show
End Sub

