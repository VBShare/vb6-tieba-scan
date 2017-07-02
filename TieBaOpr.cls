VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TieBaOpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary
'@@require
'@@reference
'  Microsoft WinHTTP Services, version 5.1
'  Microsoft Script Control 1.0
'  CHashTable
'  CList
'  VB_MD5
'@@license
'@@author
'@@create
'@@modify
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       接口继承
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有常量
'------------------------------------------------------------------------------



'------------------------------------------------------------------------------
'       公有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有API
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       事件声明
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       私有常量
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有变量
'------------------------------------------------------------------------------
Private mWinHttpReq As WinHttp.WinHttpRequest '声明一个对象
Private jsExec As New ScriptControl
Public cookies As New CHashTable


Private vcodestr As String

'------------------------------------------------------------------------------
'       属性变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有API
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      类
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       初始化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
  jsExec.language = "JavaScript"
  Set mWinHttpReq = New WinHttp.WinHttpRequest
End Sub


'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
  Set jsExec = Nothing
  Set mWinHttpReq = Nothing
End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'//      事件处理
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Private Function BytesToBstr(strBody, CodeBase) '编码转换("UTF-8"或者"GB2312"或者"GBK")
  Dim ObjStream
  Set ObjStream = CreateObject("Adodb.Stream")
  With ObjStream
    .Type = 1
    .Mode = 3
    .Open
    .Write strBody
    .Position = 0
    .Type = 2
    .charset = CodeBase
    BytesToBstr = .ReadText
    .Close
  End With
  Set ObjStream = Nothing
End Function



Private Sub collectCookies(ByVal Headers As String, ByRef cookie As CHashTable)
  Dim headerLines() As String, header As String, tmp As String, keyValue() As String
  Dim i As Long, key As String, value As String
  
  headerLines = Split(Headers, vbCrLf)
  
  For i = 0 To UBound(headerLines)
    header = headerLines(i)
    If InStr(1, LCase(header), "set-cookie") > 0 Then
      tmp = Split(header, ";")(0) ' 得到：Set-Cookie: clientuin=656972278
      tmp = Replace(tmp, "set-cookie:", "", , , vbTextCompare) ' 得到： clientuin=656972278
      tmp = Trim(tmp)
      keyValue = Split(tmp, "=")
      key = keyValue(0)
      value = keyValue(1)
      If cookie.ExistKey(key) Then
        If Len(value) > 0 Then
          cookie.Item(key) = value
        End If
      Else
        cookie.Item(key) = value
      End If
    End If
  Next i
End Sub

'//////////////////////////////////////////////////////////////////////////////
'//
'//      继承实现
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Public Function UTF8_URLEncoding(ByVal szInput As String)
  Dim wch As String, uch As String, szRet As String
  Dim X As Long
  Dim nAsc As Long, nAsc2 As Long, nAsc3 As Long
  If szInput = "" Then
      UTF8_URLEncoding = szInput
      Exit Function
  End If
  For X = 1 To Len(szInput)
    wch = Mid(szInput, X, 1)
    nAsc = AscW(wch)
    
    If nAsc < 0 Then nAsc = nAsc + 65536
    
    If (nAsc And &HFF80) = 0 Then
      szRet = szRet & wch
    Else
      If (nAsc And &HF000) = 0 Then
        uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
        szRet = szRet & uch
      Else
        uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
        Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
        Hex(nAsc And &H3F Or &H80)
        szRet = szRet & uch
      End If
    End If
  Next
  UTF8_URLEncoding = szRet
End Function
Public Function SendRequest(ByVal url As String, ByVal vcode As String, ByVal cookie As String)
  On Error GoTo FuckError
  Dim responseBody As String, posturl As String, data As String, pid As String
  pid = getpid(url)
  
  data = "ie=utf-8&tbs=eee4bf3814aae4b41497632594&complaint_type=spam&vcode_id=" & vcodestr & "&vcode=" & UTF8_URLEncoding(vcode) & "&tid=" & Mid(url, InStrRev(url, "/") + 1) & "&pid=" & pid & "&cid=0"

  mWinHttpReq.Open "POST", "http://tieba.baidu.com/complaint/commit", True
  mWinHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  mWinHttpReq.SetRequestHeader "Host", "tieba.baidu.com"
  mWinHttpReq.SetRequestHeader "Connection", "keep-alive"
  mWinHttpReq.SetRequestHeader "Content-Length", Len(data)
  mWinHttpReq.SetRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
  mWinHttpReq.SetRequestHeader "Origin", "http://tieba.baidu.com"
  mWinHttpReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.130 Safari/537.36"
  mWinHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  mWinHttpReq.SetRequestHeader "Referer", url
  mWinHttpReq.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
  mWinHttpReq.SetRequestHeader "Cookie", frmConfig.txtCookie.Text
  mWinHttpReq.Send data       '发送
  mWinHttpReq.WaitForResponse  '异步发送
  responseBody = BytesToBstr(mWinHttpReq.responseBody, "UTF-8")
  SendRequest = responseBody
  Exit Function
FuckError:
  Err.Clear
End Function


Public Sub GetVarifyCode(ByVal cookie As String, Optional ByVal DestFile As String = "")
  On Error GoTo FuckError
  Dim responseBody As String, url As String, ptLocalTk As String
  
  Dim image() As Byte
  
  If DestFile = "" Then
    DestFile = App.Path & "\vcode.png"
  End If
  url = "http://tieba.baidu.com/complaint/vcode?t=1497620939399"
  mWinHttpReq.Open "GET", url, True
  mWinHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  mWinHttpReq.SetRequestHeader "Host", "tieba.baidu.com"
  mWinHttpReq.SetRequestHeader "Connection", "Keep-Alive"
  mWinHttpReq.SetRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
  mWinHttpReq.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
  mWinHttpReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko Core/1.53.2669.400 QQBrowser/9.6.11054.400"
  'mWinHttpReq.SetRequestHeader "Referer", "https://xui.ptlogin2.qq.com/cgi-bin/xlogin?appid=11000101&target=self&s_url=https%3A%2F%2Fpay.qq.com%2Fipay%2Flogin-proxy.html"
  mWinHttpReq.SetRequestHeader "Accept-Language", "zh-Hans-CN,zh-Hans;q=0.8,en-US;q=0.5,en;q=0.3"
  mWinHttpReq.SetRequestHeader "Cookie", cookie
  mWinHttpReq.Send       '发送
  mWinHttpReq.WaitForResponse  '异步发送
  responseBody = BytesToBstr(mWinHttpReq.responseBody, "UTF-8")
  
  jsExec.Reset
  jsExec.AddCode "function a() { return " & responseBody & ";}"
  vcodestr = jsExec.Run("a").vcodestr
  
  If vcodestr = "" Then Exit Sub
    
  url = "http://tieba.baidu.com/cgi-bin/genimg?" & vcodestr

  mWinHttpReq.Open "GET", url, True
  mWinHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  mWinHttpReq.SetRequestHeader "Host", "tieba.baidu.com"
  mWinHttpReq.SetRequestHeader "Connection", "Keep-Alive"
  mWinHttpReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko Core/1.53.2669.400 QQBrowser/9.6.11054.400"
  mWinHttpReq.SetRequestHeader "Accept", "image/png, image/svg+xml, image/jxr, image/*;q=0.8, */*;q=0.5"
  'mWinHttpReq.SetRequestHeader "Referer", "http://tieba.baidu.com/complaint/info?type=2&cid=0&tid=5166454629&pid=108187425950"
  'mWinHttpReq.SetRequestHeader "Accept-Encoding", "gzip, deflate"
  mWinHttpReq.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
  mWinHttpReq.SetRequestHeader "Cookie", cookie
  mWinHttpReq.Send       '发送
  mWinHttpReq.WaitForResponse  '异步发送
  
  image = mWinHttpReq.responseBody
  Open DestFile For Binary As #1
    Put #1, , image
  Close #1
  Exit Sub
FuckError:
  Err.Clear
End Sub

Public Function getpid(ByVal url As String)
  Dim web As New WebCode
  Dim html As String, aCode As String
  
  
  html = web.GetHTMLCode(url, "utf-8")
  aCode = Split(html, "post_id&quot;:")(1)
  aCode = Split(aCode, ",")(0)

  getpid = aCode
End Function

