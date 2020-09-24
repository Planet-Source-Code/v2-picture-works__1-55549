VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function Details(Name As String, Version As String, Optional Description As String = "", Optional PluginType As Integer = 0, Optional GroupName As String = "")
    Name = "JED_Encode"
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Description = "JPEG Image (*.jpg,*.jpeg)|*.jpg"
    PluginType = 10
    GroupName = "Encoder"
End Function


Function About()
    MsgBox "JED - JPEG Encoder Plugin v " & App.Major & "." & App.Minor & "." & App.Revision & " For Picture Works" & vbCr & _
            "Use ""JPEG Encoder Class"" 0.9 beta  [26 / November / 2003] Written By John Korejwa  <korejwa@tiac.net>", vbInformation, "JED - JPEG Encoder Plugin For Picture Works"
End Function


