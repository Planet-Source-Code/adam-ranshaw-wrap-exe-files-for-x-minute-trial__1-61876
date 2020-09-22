Attribute VB_Name = "reg"
Option Explicit

Private Const appname As String = "TrialServer"
Private Const section As String = "paths"


Public Sub Save_Settings(stKey As String, value As String)
  SaveSetting appname, section, stKey, value
End Sub
Public Function Get_Settings(stKey As String, Optional value_def As String = "") As String
  Get_Settings = GetSetting(appname, section, stKey, value_def)
End Function

