Attribute VB_Name = "basPlugins"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Enum PlugType
    Unknown = -1
    Filter = 0
    Encoder = 10
    Decoder = 11
    ToolBox = 20
    MenuBar = 21
End Enum

Type PlugDetails
    Name As String
    Version As String
    Description As String
    PluginType As PlugType
    GroupName As String
End Type
Public Type Plug
    objMain As Object
    objInfo As Object
    objDetails As PlugDetails
    Name As String
End Type
Global sFileName As String
Global Plugs() As Plug
Global objForm() As Form
Global nTotalForms As Long
Global TotalPlugs As Long

Global FilterCount As Integer
Global ToolBoxCount As Integer
Global MenuBarCount As Integer
Global EncoderCount As Integer
Global DecoderCount As Integer

Global EncoderList() As String
Global DecoderList() As String

Public Const CMD_READONLY = &H1
Public Const CMD_OVERWRITEPROMPT = &H2
Public Const CMD_HIDEREADONLY = &H4
Public Const CMD_NOCHANGEDIR = &H8
Public Const CMD_SHOWHELP = &H10
Public Const CMD_ENABLEHOOK = &H20
Public Const CMD_ENABLETEMPLATE = &H40
Public Const CMD_ENABLETEMPLATEHANDLE = &H80
Public Const CMD_NOVALIDATE = &H100
Public Const CMD_ALLOWMULTISELECT = &H200
Public Const CMD_EXTENSIONDIFFERENT = &H400
Public Const CMD_PATHMUSTEXIST = &H800
Public Const CMD_FILEMUSTEXIST = &H1000
Public Const CMD_CREATEPROMPT = &H2000
Public Const CMD_SHAREAWARE = &H4000
Public Const CMD_NOREADONLYRETURN = &H8000
Public Const CMD_NOTESTFILECREATE = &H10000
Public Const CMD_NONETWORKBUTTON = &H20000
Public Const CMD_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Public Const CMD_EXPLORER = &H80000                         '  new look commdlg
Public Const CMD_NODEREFERENCELINKS = &H100000
Public Const CMD_LONGNAMES = &H200000                       '  force long names for 3.x modules

Public Const CMD_SHAREFALLTHROUGH = 2
Public Const CMD_SHARENOWARN = 1
Public Const CMD_SHAREWARN = 0


Function EnumPlugins() As Boolean
On Error Resume Next
Dim nCtr As Long
Dim PlugName As String
Dim cmdLine As String

Dim sName As String
Dim sVersion As String
Dim sDesc As String
Dim nType As Integer
Dim sGroup As String

frmPic.PlugFile.Path = App.Path & "\Plugins\"
frmPic.PlugFile.Pattern = "*.PWP"
frmPic.PlugFile.Refresh
TotalPlugs = frmPic.PlugFile.ListCount

If TotalPlugs <= 0 Then Exit Function
ReDim Preserve Plugs(TotalPlugs - 1) As Plug

'Register All Available Plugins To Avoid Any Error
For nCtr = 0 To TotalPlugs - 1
    DoEvents
    PlugName = Mid(frmPic.PlugFile.List(nCtr), 1, Len(frmPic.PlugFile.List(nCtr)) - 4)
    frmSplash.lblStatus = "Registering " & PlugName & "..."
    cmdLine = """" & App.Path & "\plugins\" & PlugName & ".pwp"" /s"
    ShellExecute 0, "OPEN", "regsvr32", cmdLine, App.Path & "\plugins\", 0
Next

ReDim EncoderList(0) As String
ReDim DecoderList(0) As String

EncoderList(0) = "Bitmap Images (*.bmp)|*.bmp"
EncoderCount = 1

DecoderList(0) = "All Picture Files (*.bmp;*.jpg;*.gif;*.cur;*.wmf)|*.bmp;*.jpg;*.gif;*.cur;*.wmf"
DecoderCount = 1

For nCtr = 0 To TotalPlugs - 1
    PlugName = Mid(frmPic.PlugFile.List(nCtr), 1, Len(frmPic.PlugFile.List(nCtr)) - 4)
    DoEvents
    frmSplash.lblStatus = "Loading " & PlugName & "..."
    Set Plugs(nCtr).objMain = CreateObject(PlugName & ".Main")
    Set Plugs(nCtr).objInfo = CreateObject(PlugName & ".Info")
    
    Plugs(nCtr).objInfo.Details sName, sVersion, sDesc, nType, sGroup
    Plugs(nCtr).objDetails.Name = sName
    Plugs(nCtr).objDetails.Version = sVersion
    Plugs(nCtr).objDetails.Description = sDesc
    Plugs(nCtr).objDetails.PluginType = nType
    Plugs(nCtr).objDetails.GroupName = sGroup
    
    
    Plugs(nCtr).Name = PlugName
    
    Load frmMain.mnuAboutPlug(frmMain.mnuAboutPlug.Count)
    frmMain.mnuAboutPlug(frmMain.mnuAboutPlug.Count - 1).Caption = PlugName
    frmMain.mnuAboutPlug(frmMain.mnuAboutPlug.Count - 1).Enabled = True
    frmMain.mnuAboutPlug(frmMain.mnuAboutPlug.Count - 1).Visible = True
    
    Select Case nType
    Case Filter
        Load frmMain.mnuPlug(frmMain.mnuPlug.Count)
        frmMain.mnuPlug(frmMain.mnuPlug.Count - 1).Caption = "&" & FilterCount + 1 & ") " & PlugName
        frmMain.mnuPlug(frmMain.mnuPlug.Count - 1).Enabled = True
        frmMain.mnuPlug(frmMain.mnuPlug.Count - 1).Visible = True
        FilterCount = FilterCount + 1
        With frmPlugManager.lstFilters
            .ListItems.Add , , sName
            .ListItems(.ListItems.Count).SubItems(1) = sVersion
            .ListItems(.ListItems.Count).SubItems(2) = sDesc
        End With
    Case Encoder
        ReDim Preserve EncoderList(EncoderCount) As String
        EncoderList(EncoderCount) = sDesc
        EncoderCount = EncoderCount + 1
        
        With frmPlugManager.lstEncoders
            .ListItems.Add , , sName
            .ListItems(.ListItems.Count).SubItems(1) = sVersion
            .ListItems(.ListItems.Count).SubItems(2) = sDesc
        End With
        
    Case Decoder
        ReDim Preserve DecoderList(DecoderCount) As String
        DecoderList(DecoderCount) = sDesc
        DecoderCount = DecoderCount + 1
        
        With frmPlugManager.lstDecoders
            .ListItems.Add , , sName
            .ListItems(.ListItems.Count).SubItems(1) = sVersion
            .ListItems(.ListItems.Count).SubItems(2) = sDesc
        End With
        
    Case ToolBox
        ToolBoxCount = ToolBoxCount + 1
        With frmPlugManager.lstToolBox
            .ListItems.Add , , sName
            .ListItems(.ListItems.Count).SubItems(1) = sVersion
            .ListItems(.ListItems.Count).SubItems(2) = sDesc
        End With
        
    Case MenuBar
        MenuBarCount = MenuBarCount + 1
        With frmPlugManager.lstMenuBar
            .ListItems.Add , , sName
            .ListItems(.ListItems.Count).SubItems(1) = sVersion
            .ListItems(.ListItems.Count).SubItems(2) = sDesc
        End With
    End Select
    
Next
    frmMain.mnuPlug(0).Visible = False
    frmMain.mnuAboutPlug(0).Visible = False
    frmSplash.lblStatus = "V2 Softwares <v2softwares@yahoo.com>"
End Function




Public Function GetEncoder(nIndex As Integer) As Integer
Dim nCtr As Integer

For nCtr = 0 To TotalPlugs
        If Plugs(nCtr).objDetails.Description = EncoderList(nIndex - 1) Then
            GetEncoder = nCtr
            Exit Function
        End If
Next nCtr
GetEncoder = -1
End Function


Public Function GetDecoder(nIndex As Integer) As Integer
Dim nCtr As Integer

For nCtr = 0 To TotalPlugs
    If Plugs(nCtr).objDetails.Description = DecoderList(nIndex - 1) Then
        GetDecoder = nCtr
        Exit Function
    End If
Next nCtr
GetDecoder = -1
End Function

Public Function GetFilter(sName As String) As Integer
Dim nCtr As Integer
Dim nCtr1 As Integer

For nCtr = 0 To TotalPlugs - 1
        If Plugs(nCtr).objDetails.Name = sName Then
            GetFilter = nCtr
            Exit Function
        End If
Next nCtr
GetFilter = -1
End Function


Public Function GetEncoderList() As String
Dim nCtr As Integer
Dim sTmp As String

For nCtr = 0 To EncoderCount - 1
    sTmp = sTmp & EncoderList(nCtr) & "|"
Next nCtr
GetEncoderList = sTmp
End Function


Public Function GetDecoderList() As String
Dim nCtr As Integer
Dim sTmp As String

For nCtr = 0 To DecoderCount - 1
    sTmp = sTmp & DecoderList(nCtr) & "|"
Next nCtr
GetDecoderList = sTmp
End Function


