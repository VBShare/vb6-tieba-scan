Attribute VB_Name = "StartUp"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public eScanLog As New DBScanLog
Public eTieba As New DBTieBa
Public ht As New CHashTable
Public tb As New CHashTable

Public Sub OpenWeb(ByVal URLs As String) 'ok at 12-05-10[RE]
  '�����ܣ�����ҳ
  Dim lngReturn As Long
  lngReturn = ShellExecute(0, "open", URLs, "", "", 1)
End Sub

Sub Main()
  Dim DbPath As String
  DbPath = Replace(App.Path & "\tieba_scan.mdb", "\\", "\")
  If Dir(DbPath) = "" Then
    CreateDb DbPath
  End If
  
  eScanLog.InitConn DbPath
  eTieba.InitConn DbPath
  
  Load frmTieBaScan
  frmTieBaScan.Show
End Sub

Public Sub CreateDb(ByVal DbPath As String)
  Dim dbc As New DbCreateHelper
  Dim mDbScanLog As DBModel, mDbTieba As DBModel
  Set mDbScanLog = New DBScanLog
  Set mDbTieba = New DBTieBa
  
  dbc.SetDbFile DbPath
  dbc.InitDbFromModels mDbTieba, mDbScanLog
End Sub