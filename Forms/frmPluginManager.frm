VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPlugManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plugins Manager"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   Icon            =   "frmPluginManager.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstMenuBar 
      Height          =   4875
      Left            =   180
      TabIndex        =   5
      ToolTipText     =   "Lists Available MenuBars."
      Top             =   540
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8599
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.ListView lstToolBox 
      Height          =   4875
      Left            =   180
      TabIndex        =   4
      ToolTipText     =   "Lists Available ToolBoxes."
      Top             =   540
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8599
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstDecoders 
      Height          =   4875
      Left            =   180
      TabIndex        =   3
      ToolTipText     =   "Lists Available Decoders."
      Top             =   540
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8599
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.ListView lstEncoders 
      Height          =   4875
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "Lists Available Encoders."
      Top             =   540
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8599
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.ListView lstFilters 
      Height          =   4875
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "Lists Available Filters."
      Top             =   540
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8599
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5445
      Left            =   200
      TabIndex        =   0
      Top             =   90
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9604
      MultiRow        =   -1  'True
      Style           =   1
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filters"
            Key             =   "FILTER"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Encoders"
            Key             =   "ENCODER"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Decoders"
            Key             =   "DECODER"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ToolBox"
            Key             =   "TOOLBOX"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MenuBar"
            Key             =   "MENUBAR"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   195
      TabIndex        =   6
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmPlugManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Me.Hide
End Sub

Private Sub lstDecoders_DblClick()
    Call Plugs(GetDecoder(lstDecoders.SelectedItem.index + 1)).objInfo.About
End Sub

Private Sub lstDecoders_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuAbout.Caption = "&About " & lstDecoders.SelectedItem.Text
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub lstEncoders_DblClick()
    Call Plugs(GetEncoder(lstEncoders.SelectedItem.index + 1)).objInfo.About
End Sub

Private Sub lstEncoders_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuAbout.Caption = "&About " & lstEncoders.SelectedItem.Text
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub lstFilters_DblClick()
    Call Plugs(GetFilter(lstFilters.SelectedItem.Text)).objInfo.About
End Sub



Private Sub lstFilters_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuAbout.Caption = "&About " & lstFilters.SelectedItem.Text
        PopupMenu mnuPopUp
    End If
End Sub





Private Sub mnuAbout_Click()
    Select Case TabStrip.SelectedItem.index
        Case 1
            lstFilters_DblClick
        Case 2
            mnuAbout.Caption = "&About " & lstEncoders.SelectedItem.Text
            lstEncoders_DblClick
        Case 3
            mnuAbout.Caption = "&About " & lstDecoders.SelectedItem.Text
            lstDecoders_DblClick
        Case 4
            mnuAbout.Caption = "&About " & lstToolBox.SelectedItem.Text
        Case 5
            mnuAbout.Caption = "&About " & lstMenuBar.SelectedItem.Text
    End Select
End Sub

Private Sub TabStrip_Click()
    Select Case TabStrip.SelectedItem.index
        Case 1              'Filters
            lstFilters.Visible = True
            lstEncoders.Visible = False
            lstDecoders.Visible = False
            lstToolBox.Visible = False
            lstMenuBar.Visible = False
            lblStatus = FilterCount & " Filter(s) Loaded."
        Case 2              'Encoders
            lstFilters.Visible = False
            lstEncoders.Visible = True
            lstDecoders.Visible = False
            lstToolBox.Visible = False
            lstMenuBar.Visible = False
            lblStatus = EncoderCount & " Encoder(s) Loaded."
       
        Case 3              'Decoders
            lstFilters.Visible = False
            lstEncoders.Visible = False
            lstDecoders.Visible = True
            lstToolBox.Visible = False
            lstMenuBar.Visible = False
            lblStatus = DecoderCount & " Decoder(s) Loaded."
        
        Case 4              'ToolBox
            lstFilters.Visible = False
            lstEncoders.Visible = False
            lstDecoders.Visible = False
            lstToolBox.Visible = True
            lstMenuBar.Visible = False
            lblStatus = ToolBoxCount & " ToolBox(es) Loaded."
        
        Case 5              'MenuBar
            lstFilters.Visible = False
            lstEncoders.Visible = False
            lstDecoders.Visible = False
            lstToolBox.Visible = False
            lstMenuBar.Visible = True
            lblStatus = MenuBarCount & " MenuBar(s) Loaded."
        
    End Select
End Sub
