VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBScanLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements DBModel
'Private
Private m_TableName As String
Private m_TableFields As String
'Private m_DBH As New AdodbHelper
Public m_DBH As New clsAccessHelper

Private m_UserName As String
Private m_IsChain As Boolean


Private Sub Class_Initialize()
  m_IsChain = False
  m_TableName = "scan_logs"
  m_TableFields = "id:integer,bar_id:integer,keyword:string,title:text,url:text,reported:text"
End Sub

Private Property Get DbModel_Db() As clsAccessHelper
  Set DbModel_Db = m_DBH
End Property
Public Property Get Db() As clsAccessHelper
  Set Db = DbModel_Db
End Property

Private Sub DBModel_InitConn(ByVal dbFilePath As String)
  m_DBH.SetConnToFile dbFilePath
End Sub
Public Sub InitConn(ByVal dbFilePath As String)
  DBModel_InitConn dbFilePath
End Sub

Private Property Get DbModel_TableFields() As String
  DbModel_TableFields = m_TableFields
End Property
Public Property Get TableFields() As String
  TableFields = DbModel_TableFields
End Property

Private Property Get DbModel_TableName() As String
  DbModel_TableName = m_TableName
End Property
Public Property Get TableName() As String
  TableName = DbModel_TableName
End Property

Private Function DBModel_Where(ByVal Conditions As String, ParamArray Params() As Variant) As ADODB.Recordset
  Dim SQL As String
  If Len(Conditions) = 0 Then
    SQL = "select * from " & m_TableName
  Else
    SQL = "select * from " & m_TableName & " where " & Conditions
  End If
  Set DBModel_Where = m_DBH.ExecParamQuery(SQL, Params)
End Function

Public Function Where(ByVal Conditions As String, ParamArray Params() As Variant) As ADODB.Recordset
  Dim SQL As String
  If Len(Conditions) = 0 Then
    SQL = "select * from " & m_TableName
  Else
    SQL = "select * from " & m_TableName & " where " & Conditions
  End If
  Set Where = m_DBH.ExecParamQuery(SQL, Params)
End Function

Public Sub Create(ByVal bar_id As Long, ByVal keyword As String, ByVal url As String, ByVal title As String)
  Dim SQL As String
  Dim id As Long
  Dim Res As ADODB.Recordset
  Set Res = m_DBH.ExecQuery("select id from " & m_TableName & " order by id desc")
  If Res.RecordCount = 0 Then
    id = 0
  Else
    id = Res.Fields("id").Value
  End If
  id = id + 1
  '"id:integer,bar_id:integer,keyword:string,url:text"
  SQL = "INSERT INTO `" & m_TableName & "` (`id`, `bar_id`, `keyword`, `url`, `title`) VALUES (" & _
        "?, ?, ?, ?, ?);"
  m_DBH.ExecParamNonQuery SQL, id, bar_id, keyword, url, title
End Sub

Public Function RecordNumber() As Long
  Dim SQL As String
  SQL = "SELECT COUNT(1) FROM " & m_TableName
  RecordNumber = CLng(m_DBH.ExecScalar(SQL))
End Function
