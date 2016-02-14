VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ArticleEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文章管理"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12435
   Icon            =   "ArticleEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   12435
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List2 
      Height          =   240
      Left            =   1800
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   6000
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "文章浏览"
      Height          =   6855
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.CommandButton Command2 
         Caption         =   "访问来源(&V)"
         Height          =   375
         Left            =   7200
         TabIndex        =   21
         Top             =   6360
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox contents 
         Height          =   5295
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9340
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"ArticleEdit.frx":08CA
      End
      Begin VB.ComboBox section 
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "转到(&Conv)"
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox title 
         Height          =   270
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   8655
      End
      Begin VB.TextBox adder 
         Height          =   270
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox addtime 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton DeleteA 
         Caption         =   "删除(&Del)"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "分类"
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "标题"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "添加者"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "添加日期"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "分类"
         Height          =   255
         Left            =   4800
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.Label fenlei 
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "文章号"
         Height          =   255
         Left            =   7080
         TabIndex        =   6
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
         Left            =   7680
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Label Label10 
      Caption         =   "分类"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "文章列表"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "ArticleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function ArticleById(ByVal IdNum As String) 'ok at 11-10-29
    Set art = eDocument.Where("`IdNum` = ?", IdNum)
    
    If art.RecordCount = 0 Then
      MsgBox "未知错误！"
      adh.ReleaseRecordset art
      Exit Function
    End If

    art.MoveFirst

    'load to boxs
    With art
        title.Text = CNull(.fields("Topic"))
        adder.Text = CNull(.fields("auser"))
        addtime.Text = CNull(.fields("AddTime"))
        fenlei.Caption = CNull(.fields("Class"))
        no.Caption = CNull(.fields("IdNum"))
        contents.TextRTF = .fields("Content")
    End With

    adh.ReleaseRecordset art
End Function

Function RefreshClass() 'ok at 11-10-29
    Combo1.Clear
    section.Clear
    Set res = eClassOf.Where("`userName` = ?", nowLogin)

    If res.RecordCount = 0 Then
      adh.ReleaseRecordset res
      Combo1.AddItem "新增分类"
      Exit Function
    End If

    Do While Not res.EOF = True
        section.AddItem CNull(res.fields("className"))
        Combo1.AddItem CNull(res.fields("className"))
        res.MoveNext
    Loop

    Combo1.AddItem "新增分类"
    adh.ReleaseRecordset res
End Function

Function LoadArticle(ByVal user As String) 'ok at 11-10-29fist run
    '清理
    List1.Clear
    List2.Clear
    title.Text = ""
    adder.Text = ""
    addtime.Text = ""
    fenlei.Caption = ""
    contents.Text = ""
    '添加
    Set art = eDocument.Where("`auser` = ?", user)

    If art.RecordCount = 0 Then
      adh.ReleaseRecordset art
      Exit Function
    End If
    art.MoveFirst

    Do While Not art.EOF = True
        List1.AddItem art.fields("Topic")
        List2.AddItem art.fields("IdNum")
        art.MoveNext
    Loop

    adh.ReleaseRecordset art
End Function

Private Sub Combo1_Click() 'ok at 11-12-29

    Dim newclass As String

    If Combo1.Text = "新增分类" Then
        newclass = InputBox("请输入新分类名", "新分类")

        If newclass = "新增分类" Or newclass = "" Then Combo1.ListIndex = 0: Exit Sub
        Call AddClass(ReWind(newclass), nowLogin)
        Call Service("RefreshClass")
    End If

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
End Sub

Private Sub Command1_Click() '转移分类ok at 11-10-29

    If section.Text = "" Then Exit Sub
    Set art = eDocument.Where("`IdNum` = ?", no.Caption)

    If art.RecordCount = 0 Then
       MsgBox "失败！", , "记录不存在"
       adh.ReleaseRecordset art
       Exit Sub
    End If
    adh.ReleaseRecordset art
    eDocument.TransferToClass section.Text, no.Caption
    MsgBox "分类转移成功！", vbInformation, "成功"
End Sub

Private Sub Command2_Click()

  Dim url As String
  Dim IdNum As String
  IdNum = no.Caption
  If IdNum = "" Then Exit Sub

  Set res = eDocument.Where("`IdNum` = ?", no.Caption)

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

Private Sub DeleteA_Click() 'ok at 11-10-28
  If no.Caption = "" Then Exit Sub
  'delete
  eDocument.Delete no.Caption
  'reset form
  title.Text = ""
  adder.Text = ""
  addtime.Text = ""
  fenlei.Caption = ""
  no.Caption = ""
  section.Clear
  Call Service("RefreshClass")
  Call Service("RefreshDocuments")
End Sub

Private Sub List1_Click()
  Call ArticleById(List2.List(List1.ListIndex))
End Sub

Function ClsArticle()
  title.Text = ""
  adder.Text = ""
  addtime.Text = ""
  no.Caption = ""
  contents.Text = ""
  Remark.Text = ""
End Function
