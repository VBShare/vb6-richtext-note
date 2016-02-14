VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Article 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文章管理大师 1.0.1.7"
   ClientHeight    =   9585
   ClientLeft      =   2745
   ClientTop       =   2070
   ClientWidth     =   14730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   14730
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton AutoTextConv 
      Caption         =   "自动转换文字"
      Height          =   375
      Left            =   6240
      TabIndex        =   39
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   10200
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "关键字与分类"
      Height          =   375
      Left            =   4560
      TabIndex        =   35
      Top             =   120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "搜文章"
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消自登录"
      Height          =   375
      Left            =   3120
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "文章阅读"
      Height          =   8895
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   10455
      Begin VB.CommandButton Command7 
         Caption         =   "访问来源"
         Height          =   405
         Left            =   8085
         TabIndex        =   38
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton Command6 
         Caption         =   "粘贴图像"
         Height          =   375
         Left            =   6720
         TabIndex        =   36
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   34
         Top             =   3240
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton scolor 
         Caption         =   "颜色"
         Height          =   375
         Left            =   2040
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton discolor 
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton bolds 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   27
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton italy 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   495
      End
      Begin RichTextLib.RichTextBox contents 
         Height          =   5295
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   9340
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":0CCA
      End
      Begin VB.TextBox Remark 
         Height          =   1455
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   6840
         Width           =   9735
      End
      Begin VB.CommandButton LastOne 
         Caption         =   "上一篇"
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   8400
         Width           =   1335
      End
      Begin VB.CommandButton NextOne 
         Caption         =   "下一篇"
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   8400
         Width           =   1335
      End
      Begin VB.CommandButton SaveChange 
         Caption         =   "保存评论和内容"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   8400
         Width           =   1695
      End
      Begin VB.TextBox addtime 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox adder 
         Height          =   270
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox title 
         Height          =   270
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   9615
      End
      Begin VB.Label Label9 
         Caption         =   "字号"
         Height          =   255
         Left            =   4800
         TabIndex        =   32
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "字体"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "评             论"
         Height          =   1095
         Left            =   240
         TabIndex        =   20
         Top             =   7080
         Width           =   255
      End
      Begin VB.Label no 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   8280
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "文章号"
         Height          =   255
         Left            =   7560
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.Label section 
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "分类"
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "添加日期"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "添加者"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "标题"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "文章列表"
      Height          =   8055
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
      Begin VB.ListBox List1 
         Height          =   7620
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "文章管理"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton WArticle 
      Caption         =   "新的文章"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "分类"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "Article"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LastKeyWord As String
Private LastIndex As Long
Private Declare Function SendMessage _
                Lib "user32 " _
                Alias "SendMessageA " (ByVal hwnd As Long, _
                                       ByVal wMsg As Long, _
                                       ByVal wParam As Long, _
                                       lParam As Any) As Long

Private Const WM_PASTE = &H302
Function LoadArticle(ByVal user As String) 'ok at 11-10-29
'done rewrite
  '清理
  Dim k As Integer

  List1.Clear
  List2.Clear
  title.Text = ""
  adder.Text = ""
  addtime.Text = ""
  section.Caption = ""
  contents.Text = ""
  '添加所有的文章到列表
  Set art = eDocument.OfUser(user).All
  If art.RecordCount = 0 Then
    adh.ReleaseRecordset art
    Exit Function
  End If
  art.MoveFirst

  Do While Not art.EOF = True
    If k = 10 Then
        DoEvents
        k = 0
    Else
        List1.AddItem art.fields("Topic")
        List2.AddItem art.fields("IdNum")
        k = k + 1
    End If

    art.MoveNext

    DoEvents
  Loop

  adh.ReleaseRecordset art
End Function

Function ArticleById(ByVal IdNum As String)
  Set art = eDocument.FindByID(IdNum)

  If art.RecordCount = 0 Then
    adh.ReleaseRecordset art
    MsgBox "未知错误！"
    Exit Function
  End If

  art.MoveFirst

  'load to boxs
  With art
    title.Text = CNull(.fields("Topic"))
    adder.Text = CNull(.fields("auser"))
    addtime.Text = CNull(.fields("AddTime"))
    section.Caption = CNull(.fields("Class"))
    no.Caption = CNull(.fields("IdNum"))
    contents.TextRTF = .fields("Content")
    Remark.Text = CNull(.fields("Remark"))
  End With

  adh.ReleaseRecordset art
End Function

Private Sub AutoTextConv_Click()

  Dim objRTB As RichTextLib.RichTextBox

  Dim i      As Long

  Dim total  As Long

  Set objRTB = Controls.Add("RICHTEXT.RichtextCtrl.1", "rtxt")
  Set rartxt = New ADODB.Recordset
  rartxt.Open "select * from Documents where txtContent is Null", conn, adOpenStatic, adLockOptimistic

  If rartxt.RecordCount = 0 Then rartxt.Close: Exit Sub
  total = rartxt.RecordCount

  Do While Not rartxt.EOF
      i = i + 1
      objRTB.TextRTF = rartxt.fields("Content")
      rartxt.fields("txtContent") = CStr(objRTB.Text)
      rartxt.Update

      DoEvents
      sm "转换进度：" & i & "/" & total, "show"
      'DoEvents
      rartxt.MoveNext
  Loop

  rartxt.Close
  sm "", "hide"
End Sub

Private Sub bolds_Click()
  'contents.SetFocus
  If contents.SelText = "" Then contents.SetFocus: Exit Sub
  contents.SelBold = Not contents.SelBold
  contents.SetFocus
End Sub

Private Sub Combo1_Click()
  If Combo1.Text = "" Then Exit Sub
  List1.Clear
  List2.Clear
  'begin
  Set res = eDocument.Where("`Class` = ? And `auser` = ?", Combo1.Text, nowLogin)
  If res.RecordCount = 0 Then
    adh.ReleaseRecordset res
    Exit Sub
  End If

  Do While Not res.EOF = True
    List1.AddItem CNull(res.fields("Topic"))
    List2.AddItem CNull(res.fields("IdNum"))
    res.MoveNext
  Loop

  adh.ReleaseRecordset res
  List1.ListIndex = 0
  LastOne.Enabled = False
End Sub

Private Sub Combo2_Click()
  contents.SetFocus

  If contents.SelText = "" Then Exit Sub
  contents.SelFontName = Combo2.Text
End Sub

Private Sub Combo3_Change()
  contents.SetFocus

  If contents.SelText = "" Then Exit Sub
  contents.SelFontSize = Combo3.Text
End Sub

Private Sub Combo3_Click()
  contents.SetFocus
  contents.SelFontSize = Combo3.Text
End Sub

Private Sub Command1_Click() 'cancel autologin sign ok at 11-10-29
  eUser.Db.ExecParamNonQuery "UPDATE Users Set isUsed = false Where uName = ?", nowLogin
End Sub

Private Sub Command2_Click() 'ok at 11-10-29
  Call ArticleEdit.LoadArticle(nowLogin)
  Call ArticleEdit.RefreshClass
  ArticleEdit.Show 1
End Sub

Private Sub Command3_Click()
  Call SearchMe.clean
  SearchMe.Show
End Sub

Private Sub Command5_Click()
    Call Service("SectionKey")
    AKind.Show
End Sub

Private Sub Command6_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)
  Picture1.Picture = Clipboard.GetData
  Clipboard.SetData Picture1.Picture
  contents.SetFocus
  Dim paste As Object
  Set paste = CreateObject("WScript.shell")
  paste.SendKeys "^V"
  Set paste = Nothing
End Sub

Private Sub Command7_Click()

  Dim url As String

  If no.Caption = "" Then Exit Sub
  Set res = eDocument.Where("IdNum = ? and 1 = ?", no.Caption, 1)

  If res.RecordCount = 0 Then
    adh.ReleaseRecordset res
    Exit Sub
  End If

  url = res.fields("Source")
  adh.ReleaseRecordset res

  If Len(url) > 10 Then
    If LCase(Mid(url, 1, 4)) = "http" Then
      Call OpenURL(url)

      Exit Sub

    End If
  End If

  MsgBox "数据非网页地址，内容如下：" & vbCrLf & url
End Sub

Private Sub contents_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim s As String

    If KeyCode = vbKeyF And (Shift And vbCtrlMask) Then
        s = InputBox("请输入搜索内容", "搜索", LastKeyWord)

        If s = "" Then Exit Sub
        
        If s = LastKeyWord And LastIndex > 0 Then
            If LastIndex < Len(contents.Text) Then
                LastIndex = MyFind(contents, LastIndex, s)
            Else
                LastIndex = MyFind(contents, 0, s)
            End If
        Else
            LastIndex = MyFind(contents, 0, s)
        End If
        
        LastKeyWord = s
        KeyCode = 0

    End If

End Sub
Function MyFind(ByRef RTF As RichTextBox, ByVal CurrentIndex As Long, ByVal searchS As String) As Long
    Dim i As Long, OriginIndex As Long
    Dim BeginPoint As Long
    Dim offset As Long
    OriginIndex = CurrentIndex
    'begin condition
    Do While BeginPoint < LenB(RTF.Text) And CurrentIndex = OriginIndex
        BeginPoint = offset + CurrentIndex + LenB(searchS) + 1
        If BeginPoint > LenB(RTF.Text) Then
            MyFind = -1
            Exit Function
        End If
        CurrentIndex = RTF.Find(searchS, BeginPoint)
        If CurrentIndex > OriginIndex Then
            MyFind = CurrentIndex
            Exit Function
        End If
        offset = offset + Len(searchS)
    Loop
    MyFind = -1
End Function
Private Sub contents_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    On Error Resume Next

    If contents.SelText = "" Then Exit Sub
    '此处加上将字体和字号显示上去的功能
    Call Modify(Combo2, contents.SelFontName)
    Call Modify(Combo3, contents.SelFontSize)
End Sub

Private Sub discolor_Click()
    contents.SetFocus

    If contents.SelText = "" Then Exit Sub
    contents.SelColor = discolor.BackColor
End Sub

Private Sub Form_Load()
  Set eDocument = eDocument.OfUser(nowLogin)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseTable

    End

End Sub

Private Sub italy_Click()
    contents.SetFocus

    If contents.SelText = "" Then Exit Sub
    contents.SelItalic = Not contents.SelItalic
    contents.SetFocus
End Sub

Private Sub LastOne_Click()
    contents.SetFocus

    If List1.ListIndex >= 1 Then
        List1.ListIndex = List1.ListIndex - 1
    End If

End Sub

Private Sub List1_Click() 'ok at 11-10-28 浏览某篇文章

    If List1.ListIndex = 0 Then LastOne.Enabled = False Else LastOne.Enabled = True
    If List1.ListIndex = List1.ListCount - 1 Then NextOne.Enabled = False Else NextOne.Enabled = True
    Call sm("正在打开……", "show")

    DoEvents
    Call ArticleById(List2.List(List1.ListIndex))

    DoEvents
    Call sm("", "hide")
End Sub

Private Sub NextOne_Click()
    contents.SetFocus

    If List1.ListIndex <= List1.ListCount - 2 Then
        List1.ListIndex = List1.ListIndex + 1
    End If

End Sub

Private Sub SaveChange_Click() 'ok at 11-10-28
    '2012-11-05 增加保存到文本内容，提供检索
    Dim sql As String
    contents.SetFocus

    Dim b() As Byte

    If Trim(no.Caption) = "" Then MsgBox "我了个去，你没有打开文章，怎么编辑？怎么保存？！！": Exit Sub
    If contents.Text = "" Then MsgBox "请填写内容", , "别坑爹了": Exit Sub
    'ready to save
    Call sm("正在保存……", "show")

    DoEvents
    b = contents.TextRTF

    sql = "Update `Documents` Set `Content` = ?, `txtContent` = ?, `Remark` = ? Where `IdNum` = ?"
    eDocument.Db.ExecParamNonQuery sql, b, contents.Text, IIf(Len(Remark.Text) > 0, Remark.Text, "暂无评论"), no.Caption

    Call sm("", "hide")
End Sub

Private Sub scolor_Click()
    contents.SetFocus
    ' 设置“取消”为True
    CommonDialog1.CancelError = True

    On Error GoTo ErrHandler

    '设置 Flags 属性
    CommonDialog1.Flags = cdlCCRGBInit
    ' 显示“颜色”对话框
    CommonDialog1.ShowColor

    ' 设置窗体的背景颜色为选定的颜色
    If contents.SelText = "" Then Exit Sub
    discolor.BackColor = CommonDialog1.Color
    contents.SelColor = discolor.BackColor

    Exit Sub
    
ErrHandler:
    ' 用户按了“取消”按钮
End Sub

Private Sub WArticle_Click() 'ok at 11-10-29
    Call NewArticle.Init
    NewArticle.Show 1, Me
End Sub

Function RefreshClass() 'ok at 11-10-28
    Combo1.Clear
    Set res = eClassOf.Db.ExecParamQuery("select `className` from `ClassOf` where `userName` = ?", nowLogin)
    If res.RecordCount = 0 Then
      adh.ReleaseRecordset res
      Exit Function
    End If

    Do While Not res.EOF = True
        Combo1.AddItem CNull(res.fields("className"))
        res.MoveNext
    Loop

    adh.ReleaseRecordset res
    
    Combo3.Clear

    For i = 1 To 72
        Combo3.AddItem i
    Next i

    For i = 0 To Screen.FontCount - 1
        Combo2.AddItem Screen.Fonts(i)
    Next i

End Function

Function sm(sMsg As String, choice As String) '显示信息

    If choice = "hide" Then
        Command4.Visible = False
    ElseIf choice = "show" Then
        Command4.Visible = True
        Command4.Caption = sMsg
    End If

End Function

Function ClsArticle()
    title.Text = ""
    adder.Text = ""
    addtime.Text = ""
    section.Caption = ""
    no.Caption = ""
    contents.Text = ""
    Remark.Text = ""
End Function

Function Modify(ByRef obj As Object, ByVal strSelect As String)

    Dim i As Long

    For i = 0 To obj.ListCount

        If obj.List(i) = strSelect Then
            obj.ListIndex = i

            Exit For

        End If

    Next i

End Function
