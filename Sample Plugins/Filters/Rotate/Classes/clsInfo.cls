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
    Name = "Rotate"
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Description = "Rotate Plugin For Picture Works"
    PluginType = 0
    GroupName = "Adjustment"
End Function


Public Function About()
    MsgBox "Rotate Plugin v " & App.Major & "." & App.Minor & "." & App.Revision & " For Picture Works By V2 Softwares.", vbInformation, "Rotate For Picture Works"
End Function

