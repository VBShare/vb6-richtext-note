VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form NewArticle 
   Caption         =   "新的文章"
   ClientHeight    =   7110
   ClientLeft      =   9960
   ClientTop       =   1365
   ClientWidth     =   10695
   Icon            =   "NewArticle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10695
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7440
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "新建文章"
      Height          =   6855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton Command2 
         Caption         =   "关闭自动"
         Height          =   375
         Left            =   6720
         TabIndex        =   26
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         TabIndex        =   25
         Top             =   2880
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton scolor 
         Caption         =   "颜色"
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton discolor 
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1320
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
         Left            =   720
         TabIndex        =   18
         Top             =   1320
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
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
      Begin RichTextLib.RichTextBox contents 
         Height          =   4455
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7858
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"NewArticle.frx":1272
      End
      Begin VB.CommandButton SaveMore 
         Caption         =   "保存继续(&C)"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "取消保存(&N)"
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   6360
         Width           =   1455
      End
      Begin VB.TextBox sourcein 
         Height          =   270
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   9255
      End
      Begin VB.ComboBox section 
         Height          =   300
         ItemData        =   "NewArticle.frx":130F
         Left            =   5400
         List            =   "NewArticle.frx":1311
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox title 
         Height          =   270
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   9495
      End
      Begin VB.TextBox adder 
         Height          =   270
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox addtime 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   3600
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton SaveReturn 
         Caption         =   "保存返回(&S)"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "字号"
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "字体"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "文章来源"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "标题"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "添加者"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "添加日期"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "分类"
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "文章号"
         Height          =   255
         Left            =   7560
         TabIndex        =   8
         Top             =   600
         Width           =   615
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
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Menu mnEdit 
      Caption         =   "编辑"
      Visible         =   0   'False
      Begin VB.Menu mnSelectAll 
         Caption         =   "全选"
      End
      Begin VB.Menu mnCopy 
         Caption         =   "复制"
      End
      Begin VB.Menu mnPaste 
         Caption         =   "粘贴"
      End
   End
End
Attribute VB_Name = "NewArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim clpStr As String
Private Declare Sub keybd_event _
                Lib "user32" (ByVal bVk As Byte, _
                              ByVal bScan As Byte, _
                              ByVal dwFlags As Long, _
                              ByVal dwExtraInfo As Long)

'添加一个根据内容自动识别分类的功能
Function RefreshClass() 'ok at 11-10-28

    If nowLogin = "" Then Exit Function
    section.Clear
    Set res = eClassOf.Where("`userName` = ?", nowLogin)

    If res.RecordCount = 0 Then
      adh.ReleaseRecordset res
      section.AddItem "新增分类"
      Exit Function
    End If

    Do While Not res.EOF = True
      section.AddItem CNull(res.fields("className"))
      res.MoveNext
    Loop

    section.AddItem "新增分类"
    adh.ReleaseRecordset res
    Combo2.Clear

    For i = 0 To Screen.FontCount - 1
      Combo2.AddItem Screen.Fonts(i)
    Next i

    For i = 1 To 72
      Combo3.AddItem i
    Next i
End Function

Private Sub bolds_Click()
    contents.SetFocus

    If contents.SelText = "" Then Exit Sub
    contents.SelBold = Not contents.SelBold
End Sub

Private Sub Combo2_Click()
    contents.SetFocus

    DoEvents
    contents.SelFontName = Combo2.Text
    contents.Font.Name = Combo2.Text
End Sub

Private Sub Combo3_Change()
    contents.SetFocus
    contents.SelFontSize = Combo3.Text
End Sub

Private Sub Combo3_Click()
    contents.SetFocus
    contents.SelFontSize = Combo3.Text
End Sub

Private Sub Command1_Click() 'cancel save ok at 11-10-28
    Timer1.Enabled = False
    Me.Hide
End Sub

Private Sub Command2_Click()

    If Command2.Caption = "关闭自动" Then
        Timer1.Enabled = False
        Command2.Caption = "自动"
    Else
        Timer1.Enabled = True
        Command2.Caption = "关闭自动"
    End If

End Sub

Private Sub contents_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim s As String

    If KeyCode = vbKeyF And (Shift And vbCtrlMask) Then
        s = InputBox("请输入搜索内容", "搜索")

        If s = "" Then Exit Sub
        KeyCode = 0
        contents.Find s
    End If

End Sub

Private Sub contents_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    If Button = 2 Then PopupMenu mnEdit
End Sub

Private Sub discolor_Click()

    If contents.SelText = "" Then Exit Sub
    contents.SelColor = discolor.BackColor
End Sub

Private Sub Form_Load()
    Call Init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = False
    Me.Hide
End Sub

Private Sub italy_Click()
    contents.SetFocus

    If contents.SelText = "" Then Exit Sub
    contents.SelItalic = Not contents.SelItalic
End Sub

Private Sub mnCopy_Click()
    Clipboard.Clear
    Clipboard.SetText contents.SelText
End Sub

Private Sub mnPaste_Click()
    contents.SelText = Clipboard.GetText
End Sub

Private Sub mnSelectAll_Click() 'ok at 12-02-10
    contents.SelStart = 0
    contents.SelLength = Len(contents.Text)
End Sub

Private Sub SaveMore_Click() 'ok at 11-10-30
  Dim sTitle As String
  Dim bContent() As Byte
  Dim sTxtContent As String
  Dim sAddTime As String
  Dim sSource As String
  Dim sClassName As String
  Dim sIdNum As String
  Dim sUserName As String
  
  sTitle = title.Text
  bContent = contents.TextRTF
  sTxtContent = contents.Text
  sAddTime = Format(Date, "yyyy-mm-dd")
  sSource = ReWindDoc(sourcein.Text)
  sClassName = section.Text
  sIdNum = no.Caption
  sUserName = ReWind(adder.Text)

  If sTitle = "" Then MsgBox "请填写标题": Exit Sub
  If sClassName = "新增分类" Then MsgBox "请选择或者新增分类": Exit Sub
  If sSource = "" Then MsgBox "请填写资料来源（网址或来源人）": Exit Sub
  If sTxtContent = "" Then MsgBox "请填写内容", , "别坑爹了": Exit Sub
  'ready to save
  Call sm("保存中……", "show")
  eDocument.Create stopic, bContent, sSource, sClassName, sIdNum, sUserName, sTxtContent
  Call sm("", "hide")
  'reseting
  title.Text = ""
  adder.Text = nowLogin
  addtime.Text = Format(Date, "yyyy-mm-dd")
  Call RefreshClass
  no.Caption = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
  sourcein.Text = ""
  contents.Text = ""
  title.SetFocus
  Call Article.LoadArticle(nowLogin)
End Sub

Private Sub SaveReturn_Click() 'ok at 11-10-30

  Dim sTitle As String
  Dim bContent() As Byte
  Dim sTxtContent As String
  Dim sAddTime As String
  Dim sSource As String
  Dim sClassName As String
  Dim sIdNum As String
  Dim sUserName As String
  
  sTitle = title.Text
  bContent = contents.TextRTF
  sTxtContent = contents.Text
  sAddTime = Format(Date, "yyyy-mm-dd")
  sSource = ReWindDoc(sourcein.Text)
  sClassName = section.Text
  sIdNum = no.Caption
  sUserName = ReWind(adder.Text)

  If sTitle = "" Then MsgBox "请填写标题": Exit Sub
  If sClassName = "新增分类" Then MsgBox "请选择或者新增分类": Exit Sub
  If sSource = "" Then MsgBox "请填写资料来源（网址或来源人）": Exit Sub
  If sTxtContent = "" Then MsgBox "请填写内容", , "别坑爹了": Exit Sub
  'ready to save
  DoEvents
  Call sm("保存中……", "show")
  eDocument.Create sTitle, bContent, sSource, sClassName, sIdNum, sUserName, sTxtContent
  DoEvents
  Call sm("", "hide")
  Call Article.LoadArticle(nowLogin)
  Timer1.Enabled = False
  Me.Hide
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

Private Sub section_Click() 'ok at 11-12-29

    Dim newclass As String

    If section.Text = "新增分类" Then
        newclass = InputBox("请输入新分类名", "新分类")

        If newclass = "新增分类" Or newclass = "" Then section.ListIndex = 0: Exit Sub
        Call AddClass(ReWind(newclass), nowLogin)
        section.AddItem newclass
        section.Text = newclass
        Call Service("RefreshClass")
    End If

End Sub

Function Init() 'ok at 11-10-28

    On Error Resume Next

    title.Text = ""
    adder.Text = nowLogin
    addtime.Text = Format(Date, "yyyy-mm-dd")
    Call RefreshClass
    no.Caption = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
    sourcein.Text = ""
    contents.Text = ""
    Timer1.Enabled = True
    title.SetFocus
End Function

Function sm(sMsg As String, choice As String)

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
    no.Caption = ""
    contents.Text = ""
    Remark.Text = ""
End Function

Private Sub Timer1_Timer()
    Timer1.Enabled = False

    On Error GoTo ends

    Dim s   As String

    Dim b() As Byte

    s = Clipboard.GetText
    Clipboard.Clear

    If s = "" Then
    Else

        If sourcein.Text = "" And Len(s) < 200 And InStr(1, s, "http") = 1 Then
            sourcein.Text = s
            Debug.Print "填充网址"
        ElseIf title.Text = "" And Len(s) < 100 And InStr(1, s, "http") = 0 Then
            title.Text = s
            Debug.Print "填充标题"
        ElseIf contents.Text = "" And Len(s) > 10 Then
            contents.Text = s: Call AutoSection(s)
            Debug.Print "初始化内容"
        ElseIf contents.Text <> "" And Len(s) > 0 Then
            contents.Text = contents.Text & vbCrLf & vbCrLf & s
            Debug.Print "新增内容"
            Call AutoSection(s)
        End If
    End If

    'clpStr = s
ends:

    Timer1.Enabled = True
End Sub

