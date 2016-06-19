Attribute VB_Name = "StartUp"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOP = 0
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

Public eScanLog As New DBScanLog
Public eTieba As New DBTieBa
Public ht As New CHashTable
Public tb As New CHashTable
Public config As New CConfig

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
  If Dir(App.Path & "\config.dat") <> "" Then
    config.LoadFromFile App.Path & "\config.dat"
  End If
  
  eScanLog.InitConn DbPath
  eTieba.InitConn DbPath
  
  Load frmTieBaScan
  Load frmConfig

  With frmConfig
    If VarType(config.Item("cookie")) <> 9 Then
      .txtCookie.text = config.Item("cookie")
    End If
    If VarType(config.Item("data")) <> 9 Then
      .txtData.text = config.Item("data")
    End If
  End With
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
