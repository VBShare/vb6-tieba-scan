VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "发帖机器"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5070
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2640
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "启动"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   6360
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private mWinHttpReq As WinHttp.WinHttpRequest
Private fs As New CFile
Private Sub Command1_Click()
  If Command1.Caption = "启动" Then
    Command1.Caption = "暂停"
    Timer1.Enabled = True
  Else
    Command1.Caption = "启动"
    Timer1.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  Set mWinHttpReq = New WinHttp.WinHttpRequest
End Sub

Private Sub Timer1_Timer()
  Dim result As String
  Timer1.Enabled = False
  If Command1.Caption = "暂停" Then
    '说明还可以执行请求
    result = SendMessage(Format(Now, "HH:mm:ss") & " 抵制垃圾信息，守卫贴吧！")
    'result = SendMessage("文化科技体育")
    If InStr(1, result, """err_code"":0") > 0 Then
      List1.AddItem Format(Now, "yyyy-MM-dd HH:mm:ss") & " 发帖成功！"
    Else
      List1.AddItem Format(Now, "yyyy-MM-dd HH:mm:ss") & " 发帖失败！"
    End If
    List1.ListIndex = List1.ListCount - 1
    'fs.WriteLineToTextFile App.Path & "\log.txt", Format(Now, "yyyy-MM-dd HH:mm:ss") & vbCrLf & result
    '说明还可以激活自动发帖
    If Command1.Caption = "暂停" Then Timer1.Enabled = True
  Else
    '说明不能自动发了
    Timer1.Enabled = False
  End If
End Sub

Private Function BytesToBstr(strBody, CodeBase) '编码转换("UTF-8"或者"GB2312"或者"GBK")
  Dim ObjStream
  Set ObjStream = CreateObject("Adodb.Stream")
  With ObjStream
    .Type = 1
    .Mode = 3
    .Open
    .Write strBody
    .position = 0
    .Type = 2
    .Charset = CodeBase
    BytesToBstr = .ReadText
    .Close
  End With
  Set ObjStream = Nothing
End Function

Public Function SendMessage(ByVal Message As String)
  On Error GoTo FuckError
  Dim responseBody As String, url As String, data As String
  
  url = "http://tieba.baidu.com/f/commit/post/add"
  data = "" '请自己填上

  mWinHttpReq.Open "POST", url, True
  mWinHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  mWinHttpReq.SetRequestHeader "Host", "tieba.baidu.com"
  mWinHttpReq.SetRequestHeader "Connection", "keep-alive"
  mWinHttpReq.SetRequestHeader "Content-Length", Len(data)
  mWinHttpReq.SetRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
  mWinHttpReq.SetRequestHeader "Origin", "http://tieba.baidu.com"
  mWinHttpReq.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
  mWinHttpReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.130 Safari/537.36"
  mWinHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
  mWinHttpReq.SetRequestHeader "Referer", "http://tieba.baidu.com/p/5182181430"
  mWinHttpReq.SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
  mWinHttpReq.SetRequestHeader "Cookie", "" '请自己填上
  mWinHttpReq.Send data       '发送
  mWinHttpReq.WaitForResponse  '异步发送
  responseBody = BytesToBstr(mWinHttpReq.responseBody, "UTF-8")
  SendMessage = responseBody
  Exit Function
FuckError:
  Err.Clear
  SendMessage = ""
End Function

Private Function UTF8_URLEncoding(ByVal szInput As String)
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
