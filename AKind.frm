VERSION 5.00
Begin VB.Form AKind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关键字与分类"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   Icon            =   "AKind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "删除"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除(&D)"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加(&A)"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   2220
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2220
      ItemData        =   "AKind.frx":24A2
      Left            =   120
      List            =   "AKind.frx":24A4
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "对应关键字"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "分类"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "AKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click() 'ok at 11-11-06

    Dim s As String

    List1.SetFocus

    If List1.ListIndex < 0 Then MsgBox "请选择分类": Exit Sub
    s = InputBox("请输入关键字", "关键字输入")

    If s = "" Then Exit Sub
    Call AddRelate(ReWindDoc(s), List1.List(List1.ListIndex))
    List2.AddItem ReWindDoc(s)
End Sub

Private Sub Command2_Click() 'ok at 11-11-06
    List2.SetFocus

    If List2.ListIndex < 0 And List1.ListIndex < 0 Then Exit Sub
    Call DelRelate(List2.List(List2.ListIndex), List1.List(List1.ListIndex))
    Call Service("SectionKey")
End Sub

Function AddRelate(ByVal Keyword As String, Sections As String) 'ok at 11-11-06
  If eRelate.KeywordExist(Keyword) Then
    MsgBox "相同关键字存在，请勿重复添加", vbCritical, "提醒"
    Exit Function
  Else
    eRelate.Create Keyword, Sections
  End If
End Function

Function DelRelate(ByVal Keyword As String, Section As String) 'ok at 11-11-06
  eRelate.RemoveRelate Section, Keyword
End Function

Function DelClass(ByVal ClsName As String) 'ok at 11-12-07
  eRelate.RemoveRelateOfClass ClsName
  eClassOf.RemoveClass ClsName
End Function

Private Sub Command3_Click()
    List1.SetFocus

    If List1.ListIndex < 0 Then Exit Sub
    Call DelClass(List1.List(List1.ListIndex))
    Call Service("SectionKey")
End Sub

Private Sub List1_Click() 'ok at 11-11-06

    If List1.ListCount < 0 Then Exit Sub
    List2.Clear
    Set res = eRelate.OfClass(List1.List(List1.ListIndex))
    
    If res.RecordCount = 0 Then
      adh.ReleaseRecordset res
      Exit Sub
    End If

    Do While Not res.EOF = True
      List2.AddItem res.fields("keyword")
      res.MoveNext
    Loop

    adh.ReleaseRecordset res
End Sub

Function RefreshKind() 'ok at 11-11-06
    List1.Clear
    List2.Clear
    Set res = eClassOf.All

    If res.RecordCount = 0 Then
      adh.ReleaseRecordset res
      Exit Function
    End If

    Do While Not res.EOF = True
      List1.AddItem res.fields("className")
      res.MoveNext
    Loop

    adh.ReleaseRecordset res
End Function
