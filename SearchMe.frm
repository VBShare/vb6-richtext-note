VERSION 5.00
Begin VB.Form SearchMe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "搜索文章【模糊搜索】"
   ClientHeight    =   3480
   ClientLeft      =   7965
   ClientTop       =   4635
   ClientWidth     =   5175
   Icon            =   "SearchMe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5175
   Begin VB.CommandButton Command3 
      Caption         =   "关闭(C&los)"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空(&Cl)"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "搜索(&Serh)"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "搜索条件"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.TextBox content 
         Height          =   1335
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox title 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "内  容  概  要"
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "标题"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "SearchMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Title.SetFocus

  Dim sql As String

  sql = "select * from Documents where"

  If Title = "" Then
    If Content.Text = "" Then
      MsgBox "请填写至少一项"
      Exit Sub
    Else
      sql = sql & " txtContent Like '%" & ReWindDoc(Content.Text) & "%' and auser='" & nowLogin & "'"
    End If
  Else
    If Content.Text = "" Then
      sql = sql & " Topic Like '%" & ReWindDoc(Title.Text) & "%' and auser='" & nowLogin & "'"
    Else
      sql = sql & " txtContent Like '%" & ReWindDoc(Content.Text) & "%' and Topic Like '%" & ReWindDoc(Title.Text) & "%' and auser='" & nowLogin & "'"
    End If
  End If

  Set res = documents.Db.ExecQuery(sql)
  res.Open sql, conn, 3, 3

  If res.RecordCount = 0 Then MsgBox "未搜索到相似信息！": res.Close: Exit Sub
  Article.List1.Clear
  Article.List2.Clear

  Do While Not res.EOF = True

      With Article
          .List1.AddItem res.fields("Topic")
          .List2.AddItem res.fields("IdNum")
      End With

      res.MoveNext
  Loop

  res.Close
  Article.List1.ListIndex = 0
  Article.LastOne.Enabled = False
End Sub

Private Sub Command2_Click()
    Call clean
End Sub

Private Sub Command3_Click()
    Call clean
    Me.Hide
End Sub

Function clean()

    On Error Resume Next

    Title.Text = ""
    Content.Text = ""
    Title.SetFocus
End Function
