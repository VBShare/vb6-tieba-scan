VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmTieBaScan 
   Caption         =   "扫描"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   11910
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Picture1 
      Caption         =   "Get"
      Height          =   375
      Left            =   3600
      TabIndex        =   29
      Top             =   6600
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   6120
      Width           =   3375
      ExtentX         =   5953
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command7 
      Caption         =   "清空"
      Height          =   495
      Left            =   7800
      TabIndex        =   27
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "反选"
      Height          =   495
      Left            =   6720
      TabIndex        =   26
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "全选"
      Height          =   495
      Left            =   5640
      TabIndex        =   25
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "提交"
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtVCode 
      Height          =   375
      Left            =   3600
      TabIndex        =   23
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton btnExportCopy 
      Caption         =   "导出到剪贴板"
      Height          =   495
      Left            =   9960
      TabIndex        =   22
      Top             =   4920
      Width           =   1785
   End
   Begin VB.CommandButton btnOpenAll 
      Caption         =   "批量打开"
      Height          =   495
      Left            =   8520
      TabIndex        =   20
      Top             =   4920
      Width           =   1300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "配置信息"
      Height          =   495
      Left            =   7080
      TabIndex        =   18
      Top             =   4920
      Width           =   1300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "批量删帖"
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   4920
      Width           =   1300
   End
   Begin VB.ListBox lstURL 
      Height          =   4470
      Left            =   5640
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Caption         =   "贴吧控制台"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton btnLoadReported 
         Caption         =   "读取已举报历史"
         Height          =   495
         Left            =   2880
         TabIndex        =   32
         Top             =   4560
         Width           =   2415
      End
      Begin VB.CommandButton btnLoadUnreported 
         Caption         =   "读取未举报历史"
         Height          =   495
         Left            =   2880
         TabIndex        =   31
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CommandButton btnSearchName 
         Caption         =   "搜索用户"
         Height          =   495
         Left            =   2880
         TabIndex        =   21
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "清空已被删除"
         Height          =   495
         Left            =   2880
         TabIndex        =   19
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton btnReadHistory 
         Caption         =   "读取历史记录"
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtPageEnd 
         Height          =   270
         Left            =   4560
         TabIndex        =   14
         Text            =   "10"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPageStart 
         Height          =   270
         Left            =   3480
         TabIndex        =   12
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtKeyWord 
         Height          =   270
         Left            =   3600
         TabIndex        =   10
         Top             =   270
         Width           =   1695
      End
      Begin VB.CommandButton btnScan 
         Caption         =   "扫描"
         Height          =   495
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton btnRemove 
         Caption         =   "删除"
         Height          =   400
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "增加"
         Height          =   400
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtTieBa 
         Height          =   270
         Left            =   840
         TabIndex        =   5
         Top             =   270
         Width           =   1935
      End
      Begin VB.ListBox lstTieBas 
         Height          =   4020
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "到"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "页码"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "关键词"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "新贴吧"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   5640
      TabIndex        =   30
      Top             =   6240
      Width           =   6135
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   5500
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "网址"
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   100
      Width           =   495
   End
End
Attribute VB_Name = "frmTieBaScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TGUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long

Private Res As ADODB.Recordset
Private currentBar As String
Private currentCheck As Long
Private tiebaOp As New TieBaOpr

'****************** Event Handler ******************

Private Sub btnAdd_Click()
  Dim barName As String
  
  barName = txtTieBa.text
  If barName = "" Then
    MsgBox "请输入贴吧名", , "提示"
    txtTieBa.SetFocus
    Exit Sub
  End If
  
  If eTieba.BarExist(barName) Then
    MsgBox "贴吧已存在", , "提示"
    txtTieBa.text = ""
    txtTieBa.SetFocus
    Exit Sub
  End If

  eTieba.Create barName
  
  tb.Item(barName) = eTieba.Where("bar_name = ?", barName).Fields("id").Value
  
  lstTieBas.AddItem barName
End Sub

Private Sub btnExportCopy_Click()
  Dim i As Long, url As String
  Dim msg As String
  msg = "以下帖子为垃圾信息或恶意灌水：" & vbCrLf
  If lstURL.ListCount = 0 Then Exit Sub
  For i = 0 To lstURL.ListCount - 1
    If lstURL.Selected(i) = True Then
      msg = msg & lstURL.List(i) & "  " & ht.Item(i) & vbCrLf
    End If
  Next i
  Clipboard.Clear
  Clipboard.SetText msg
End Sub

Private Sub btnLoadReported_Click()
  Dim Res As ADODB.Recordset
  
  If currentBar = "" Then Exit Sub
  
  Set ht = Nothing
  Set ht = New CHashTable
  lstURL.Clear
  
  Set Res = eScanLog.Where("`bar_id` = ? and `reported` = 'done'", tb.Item(currentBar))
  
  If Res.RecordCount = 0 Then
    eScanLog.Db.ReleaseRecordset Res
    Exit Sub
  End If
  
  Do While Not Res.EOF
    ht.Item(lstURL.ListCount) = Res.Fields("url").Value
    lstURL.AddItem Res.Fields("title").Value
    Res.MoveNext
  Loop
  
  eScanLog.Db.ReleaseRecordset Res
End Sub

Private Sub btnLoadUnreported_Click()
  Dim Res As ADODB.Recordset
  
  If currentBar = "" Then Exit Sub
  
  Set ht = Nothing
  Set ht = New CHashTable
  lstURL.Clear
  
  Set Res = eScanLog.Where("`bar_id` = ? and `reported` is null", tb.Item(currentBar))
  
  If Res.RecordCount = 0 Then
    eScanLog.Db.ReleaseRecordset Res
    Exit Sub
  End If
  
  Do While Not Res.EOF
    ht.Item(lstURL.ListCount) = Res.Fields("url").Value
    lstURL.AddItem Res.Fields("title").Value
    Res.MoveNext
  Loop
  
  eScanLog.Db.ReleaseRecordset Res
End Sub

Private Sub btnOpenAll_Click()
  Dim i As Long, url As String
  If lstURL.ListCount = 0 Then Exit Sub
  For i = 0 To lstURL.ListCount - 1
    If lstURL.Selected(i) Then
      url = ht.Item(i)
      OpenWeb url
      DoEvents
    End If
  Next i
End Sub

Private Sub btnReadHistory_Click()
  Dim Res As ADODB.Recordset
  
  If currentBar = "" Then Exit Sub
  
  Set ht = Nothing
  Set ht = New CHashTable
  lstURL.Clear
  
  Set Res = eScanLog.Where("`bar_id` = ?", tb.Item(currentBar))
  
  If Res.RecordCount = 0 Then
    eScanLog.Db.ReleaseRecordset Res
    Exit Sub
  End If
  
  Do While Not Res.EOF
    ht.Item(lstURL.ListCount) = Res.Fields("url").Value
    lstURL.AddItem Res.Fields("title").Value
    Res.MoveNext
  Loop
  
  eScanLog.Db.ReleaseRecordset Res
End Sub

Private Sub btnScan_Click()
  'page:
  'http://tieba.baidu.com/f?kw=%E5%B0%8F%E7%B1%B3&ie=utf-8&pn=50
  Dim i As Long
  Dim maxPage As Long, minPage As Long
  If currentBar = "" Then
    MsgBox "未选择目标贴吧"
    Exit Sub
  End If
  If IsNumeric(txtPageStart.text) Then
    minPage = CLng(txtPageStart.text)
  Else
    MsgBox "起始页码不为数字"
    txtPageStart.text = ""
    txtPageStart.SetFocus
    Exit Sub
  End If
  If IsNumeric(txtPageEnd.text) Then
    maxPage = CLng(txtPageEnd.text)
  Else
    MsgBox "终止页码不为数字"
    txtPageEnd.text = ""
    txtPageEnd.SetFocus
    Exit Sub
  End If
  If minPage > maxPage Then
    maxPage = minPage
    minPage = CLng(txtPageEnd.text)
  End If
  If minPage < 1 Then
    MsgBox "最小页码不得小于1"
    txtPageStart.SetFocus
    Exit Sub
  End If
  '清空
  Set ht = Nothing
  Set ht = New CHashTable
  lstURL.Clear
  '开始扫描
  For i = minPage To maxPage
    SetStatus "当前页码：" & i
    Call scanPage(currentBar, txtKeyWord.text, i)
    DoEvents
  Next i
  SetStatus "完成扫描"
End Sub

Private Sub btnSearchName_Click()
  'page:
  'http://tieba.baidu.com/f?kw=%E5%B0%8F%E7%B1%B3&ie=utf-8&pn=50
  Dim i As Long
  Dim maxPage As Long, minPage As Long
  If currentBar = "" Then
    MsgBox "未选择目标贴吧"
    Exit Sub
  End If
  If txtKeyWord.text = "" Then
    MsgBox "请设置关键字"
    txtKeyWord.SetFocus
    Exit Sub
  End If
  If IsNumeric(txtPageStart.text) Then
    minPage = CLng(txtPageStart.text)
  Else
    MsgBox "起始页码不为数字"
    txtPageStart.text = ""
    txtPageStart.SetFocus
    Exit Sub
  End If
  If IsNumeric(txtPageEnd.text) Then
    maxPage = CLng(txtPageEnd.text)
  Else
    MsgBox "终止页码不为数字"
    txtPageEnd.text = ""
    txtPageEnd.SetFocus
    Exit Sub
  End If
  If minPage > maxPage Then
    maxPage = minPage
    minPage = CLng(txtPageEnd.text)
  End If
  If minPage < 1 Then
    MsgBox "最小页码不得小于1"
    txtPageStart.SetFocus
    Exit Sub
  End If
  '清空
  Set ht = Nothing
  Set ht = New CHashTable
  lstURL.Clear
  '开始扫描
  For i = minPage To maxPage
    SetStatus "当前页码：" & i
    Call scanPagePoster(currentBar, txtKeyWord.text, i)
    DoEvents
  Next i
End Sub

Private Sub Command1_Click()
  Dim i As Long, url As String
  If lstURL.ListCount = 0 Then Exit Sub
  If frmConfig.txtCookie.text = "" Or frmConfig.txtData.text = "" Then
    MsgBox "请配置参数"
    Exit Sub
  End If
  For i = 0 To lstURL.ListCount - 1
    If lstURL.Selected(i) = True Then
      SetStatus "删帖 - " & lstURL.List(i)
      DoEvents
      url = ht.Item(i)
      Call DeleteOnePost(url)
      SetStatus "删帖完成 - " & lstURL.List(i)
      DoEvents
    End If
  Next i
  SetStatus "待命"
End Sub

Private Sub Command2_Click()
  frmConfig.Show
End Sub

Private Sub Command3_Click()
  Dim Res As ADODB.Recordset
  Dim webGet As New WebCode
  Dim url As String
  Dim total As Long, index As Long, delcount As Long
  
  If currentBar = "" Then Exit Sub
  
  Set ht = Nothing
  Set ht = New CHashTable
  
  Set Res = eScanLog.Where("`bar_id` = ?", tb.Item(currentBar))
  
  If Res.RecordCount = 0 Then
    eScanLog.Db.ReleaseRecordset Res
    Exit Sub
  End If
  total = Res.RecordCount
  index = 0
  Do While Not Res.EOF
    index = index + 1
    url = Res.Fields("url").Value
    If InStr(1, webGet.GetHTMLCode(url), "doodle-404") > 0 Then
      eScanLog.Db.ExecParamNonQuery "delete from scan_logs where id=?", Res.Fields("id").Value
      delcount = delcount + 1
    End If
    SetStatus "[" & index & "/" & total & "/" & delcount & "]检查：" & url
    Res.MoveNext
  Loop
  
  eScanLog.Db.ReleaseRecordset Res
  
  'SetStatus "待命"
End Sub

Private Sub Command4_Click()
  Dim url As String, result As String
  Dim i As Long
  If lstURL.ListCount = 0 Then Exit Sub
  For i = 0 To lstURL.ListCount - 1
    If lstURL.Selected(i) Then
      url = ht.Item(i)
      Exit For
    End If
    DoEvents
  Next i
  If url <> "" Then
    result = tiebaOp.SendRequest(url, txtVCode.text, frmConfig.txtCookie.text)
    Debug.Print result
    If InStr(1, result, """errmsg"":""success""}") > 0 Then
      Label6.Caption = Time & ": 成功举报"
      eScanLog.Db.ExecParamNonQuery "update scan_logs set `reported`='done' where `url` = ?", url
      lstURL.Selected(i) = False
    ElseIf InStr(1, result, "already complainted") > 0 Then
      Label6.Caption = Time & ": 已举报过了"
      eScanLog.Db.ExecParamNonQuery "update scan_logs set `reported`='done' where `url` = ?", url
      lstURL.Selected(i) = False
    Else
      Label6.Caption = Time & ": 未知"
    End If
  End If
  Call Picture1_Click
  txtVCode.text = ""
  txtVCode.SetFocus
End Sub

Private Sub Command5_Click()
  Dim i As Long
  If lstURL.ListCount = 0 Then Exit Sub
  For i = 0 To lstURL.ListCount - 1
    lstURL.Selected(i) = True
    DoEvents
  Next i
End Sub

Private Sub Command6_Click()
  Dim i As Long
  If lstURL.ListCount = 0 Then Exit Sub
  For i = 0 To lstURL.ListCount - 1
    lstURL.Selected(i) = Not lstURL.Selected(i)
    DoEvents
  Next i
End Sub

Private Sub Command7_Click()
  Dim i As Long
  If lstURL.ListCount = 0 Then Exit Sub
  For i = 0 To lstURL.ListCount - 1
    lstURL.Selected(i) = False
    DoEvents
  Next i
End Sub

Private Sub Form_Load()
  InitForm
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Debug.Print "end"
  End
End Sub

Private Sub lstTieBas_Click()
  currentBar = lstTieBas.List(lstTieBas.ListIndex)
End Sub


'****************** Methods ******************
Private Sub SetStatus(ByVal description As String)
  lblStatus.Caption = description
  DoEvents
End Sub

Private Sub scanPage(ByVal barName As String, ByVal keyword As String, ByVal pageIndex As Long)
  On Error Resume Next
  Dim web As New WebCode
  Dim pageCode As String
  Dim url As String, baseUrl As String
  Dim fullUrl As String, title As String
  
  Dim htmlDom As New HTMLDocument
  Dim htmlHrefs As IHTMLElementCollection
  Dim htmlHref As HTMLAnchorElement
  
  Dim i As Long
  
  baseUrl = "http://tieba.baidu.com/"

  url = baseUrl & "f?kw=" & barName & "&ie=utf-8&pn=" & (pageIndex - 1) * 50
  pageCode = web.GetHTMLCode(url, "utf-8")
  htmlDom.body.innerHTML = pageCode

  Set htmlHrefs = htmlDom.getElementsByTagName("a")
  eScanLog.m_DBH.OpenDB
  For Each htmlHref In htmlHrefs
    fullUrl = baseUrl & htmlHref.pathname
    title = htmlHref.title
    If LCase(htmlHref.ClassName) = "j_th_tit " Then
      If keyword = "" Then
        ht.Item(lstURL.ListCount) = fullUrl
        'save to db
        If eScanLog.Where("url = ?", fullUrl).RecordCount = 0 Then
          eScanLog.Create tb.Item(barName), keyword, fullUrl, title
          SetStatus "保存：" & title & " - " & fullUrl
        Else
          SetStatus "缓存：" & title & " - " & fullUrl
        End If
        lstURL.AddItem title
      Else
        If title Like "*" & keyword & "*" Then
          ht.Item(lstURL.ListCount) = fullUrl
          'save to db
          If eScanLog.Where("url = ?", fullUrl).RecordCount = 0 Then
            eScanLog.Create tb.Item(barName), keyword, fullUrl, title
            SetStatus "保存：" & title & " - " & fullUrl
          Else
            SetStatus "缓存：" & title & " - " & fullUrl
          End If
          lstURL.AddItem title
        End If
      End If
    End If
  Next
  eScanLog.m_DBH.CloseDB
End Sub

Private Sub scanPagePoster(ByVal barName As String, ByVal keyword As String, ByVal pageIndex As Long)
  'On Error Resume Next
  Dim web As New WebCode
  Dim pageCode As String
  Dim url As String, baseUrl As String
  Dim fullUrl As String, title As String
  
  Dim htmlDom As New HTMLDocument
  Dim htmlItems As IHTMLElementCollection
  Dim htmlItem As HTMLLIElement
  Dim titleName As Variant
  
  Dim i As Long
  
  baseUrl = "http://tieba.baidu.com/"

  url = baseUrl & "f?kw=" & barName & "&ie=utf-8&pn=" & (pageIndex - 1) * 50
  pageCode = web.GetHTMLCode(url, "utf-8")
  htmlDom.body.innerHTML = pageCode

  Set htmlItems = htmlDom.getElementsByTagName("li")
  eScanLog.m_DBH.OpenDB
  For Each htmlItem In htmlItems
    If htmlItem.ClassName = " j_thread_list clearfix" Then
      titleName = GetTitleName(htmlItem)
      title = titleName(0)
      fullUrl = titleName(2)
      If titleName(1) Like "*" & keyword & "*" Then
        ht.Item(lstURL.ListCount) = titleName(2)
        'save to db
        If eScanLog.Where("url = ?", fullUrl).RecordCount = 0 Then
          eScanLog.Create tb.Item(barName), keyword, fullUrl, title
          SetStatus "保存：" & title & " - " & fullUrl
        Else
          SetStatus "缓存：" & title & " - " & fullUrl
        End If
        lstURL.AddItem title
      End If
    End If
  Next
  eScanLog.m_DBH.CloseDB
End Sub

Private Function GetTitleName(li As HTMLLIElement) As Variant
  Dim titleName As String, userName As String, titleUrl As String
  Dim elements() As IHTMLElement
  Dim htmlHref As HTMLAnchorElement, baseUrl As String, htmlSpan As HTMLSpanElement
  baseUrl = "http://tieba.baidu.com/"
  For Each htmlHref In li.getElementsByTagName("a")
    If LCase(htmlHref.ClassName) = "j_th_tit " Then
      titleName = htmlHref.title
      titleUrl = baseUrl & htmlHref.pathname
      Exit For
    End If
    
  Next
  For Each htmlSpan In li.getElementsByTagName("span")
    If InStr(1, LCase(htmlSpan.ClassName), "tb_icon_author ") = 1 Then
      userName = Replace(htmlSpan.title, "主题作者: ", "")
      Exit For
    End If
    
  Next
  GetTitleName = Array(titleName, userName, titleUrl)
End Function

Private Function getElementsByClassName(Dom As IHTMLElement, ByVal ClassName As String) As IHTMLElement()
  Dim collection() As IHTMLElement
  Dim htmlNode As Variant, ObjClassName As String
  Dim Count As Long, i As Long
  Dim nodes As New CStack
  nodes.Push Dom
  Do While nodes.Count > 0
    Set htmlNode = nodes.Pop
    ObjClassName = getClassName(htmlNode)
    If InStr(1, ObjClassName, ClassName) Then
      ReDim Preserve collection(Count)
      Set collection(Count) = htmlNode
      Count = Count + 1
    End If
    For i = 1 To htmlNode.childNodes.length
      nodes.Push htmlNode.childNodes(i - 1)
    Next i
  Loop
  getElementsByClassName = collection
End Function

Private Function getClassName(obj) As String
  On Error GoTo DOIT
  getClassName = obj.ClassName
  Exit Function
DOIT:
  getClassName = vbEmpty
End Function

Private Sub InitForm()
  Set Res = eTieba.Db.ExecQuery("select `bar_name`,`id` from `tiebas`")
  If Res.RecordCount = 0 Then
    eTieba.Db.ReleaseRecordset Res
    Exit Sub
  End If
  Do While Not Res.EOF
    lstTieBas.AddItem Res.Fields("bar_name").Value
    tb.Item(Res.Fields("bar_name").Value) = Res.Fields("id").Value
    Res.MoveNext
  Loop
  eTieba.Db.ReleaseRecordset Res
End Sub

'删帖代码
Private Sub DeleteOnePost(ByVal url As String)
  Dim mWinHttpReq As New WinHttp.WinHttpRequest
  Dim shuju As String
  shuju = frmConfig.txtData.text & "tid=" & Mid(url, InStrRev(url, "/") + 1)
  mWinHttpReq.Open "POST", "http://tieba.baidu.com/f/commit/thread/delete", True
  mWinHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  mWinHttpReq.SetRequestHeader "Host", "tieba.baidu.com"
  mWinHttpReq.SetRequestHeader "Connection", "keep-alive"
  mWinHttpReq.SetRequestHeader "Content-Length", Len(shuju)
  mWinHttpReq.SetRequestHeader "Cache-Control", "max-age=0"
  mWinHttpReq.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
  mWinHttpReq.SetRequestHeader "Origin", "http://tieba.baidu.com"
  mWinHttpReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.130 Safari/537.36"
  mWinHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  mWinHttpReq.SetRequestHeader "Referer", url
  mWinHttpReq.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
  mWinHttpReq.SetRequestHeader "Cookie", frmConfig.txtCookie.text
  mWinHttpReq.Send shuju       '发送
  mWinHttpReq.WaitForResponse  '异步发送
  Set mWinHttpReq = Nothing
End Sub

Private Sub lstURL_Click()
  Debug.Print "click"
End Sub

Private Sub lstURL_DblClick()
  Dim url As String
  Debug.Print "dblclick"
  If lstURL.ListCount = 0 Then Exit Sub
  url = ht.Item(lstURL.ListIndex)
  OpenWeb url
End Sub

Private Sub Picture1_Click()
  tiebaOp.GetVarifyCode frmConfig.txtCookie.text
  Set Picture1.Picture = LoadPic(App.Path & "\vcode.png")
  'Picture2.Picture = LoadPic(App.Path & "\vcode.png")
  WebBrowser1.navigate "file:///E:/GitCode/VB6/vb6-tieba-scan/index.html"
End Sub

Private Function LoadPic(ByVal strFileName As String) As Picture
  Dim IID As TGUID
  With IID
    .Data1 = &H7BF80980
    .Data2 = &HBF32
    .Data3 = &H101A
    .Data4(0) = &H8B
    .Data4(1) = &HBB
    .Data4(2) = &H0
    .Data4(3) = &HAA
    .Data4(4) = &H0
    .Data4(5) = &H30
    .Data4(6) = &HC
    .Data4(7) = &HAB
  End With
  On Error GoTo LocalErr
  OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPic
  Exit Function
LocalErr:
  Set LoadPic = VB.LoadPicture(strFileName)
  Err.Clear
End Function

Private Sub txtVCode_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Command4_Click
  End If
End Sub
