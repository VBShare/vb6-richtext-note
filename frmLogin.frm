VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录"
   ClientHeight    =   2460
   ClientLeft      =   8325
   ClientTop       =   5730
   ClientWidth     =   4560
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1453.449
   ScaleMode       =   0  'User
   ScaleWidth      =   4281.594
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "新用户"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox autologin 
      Caption         =   "自动登录"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox remember 
      Caption         =   "记住密码"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox user 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Text            =   "<请输入用户名>"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   390
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   390
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   900
   End
   Begin VB.TextBox pass 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1125
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "文 章 数 据 库"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   60
      Width           =   4095
   End
   Begin VB.Label lblLabels 
      Caption         =   "用户名称(&U):"
      Height          =   270
      Index           =   0
      Left            =   345
      TabIndex        =   7
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "密码(&P):"
      Height          =   270
      Index           =   1
      Left            =   345
      TabIndex        =   8
      Top             =   1140
      Width           =   1080
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub autologin_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

    If autologin.Value = 0 Then remember.Value = 1
End Sub

Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    Call CloseTable

    End

End Sub

Private Sub cmdOK_Click()

    On Error Resume Next

    If user.Text = "" Or user.Text = "<请输入用户名>" Then MsgBox "请选择用户名或新建用户": user.SetFocus: Exit Sub
    If pass.Text = "" Then MsgBox "请输入您的密码": pass.SetFocus: Exit Sub

    'check exits
    If ExistUser(user.Text, pass.Text) = True Then
        nowLogin = user.Text
        LoginOk = True
        Call clearAutoSign '清理自动登录记录
        Call SaveUserSetting(nowLogin) '保存当前用户的
        Article.Show '显示界面

        DoEvents
        'Show Loading States
        Article.Command4.Caption = "正在加载文章目录……"
        Article.Command4.Visible = True

        DoEvents
        Call Article.LoadArticle(nowLogin)

        DoEvents
        Article.Command4.Caption = "正在加载文章分类……"

        DoEvents
        Call Article.RefreshClass
        Article.Command4.Visible = False
        Article.Command4.Caption = ""
        Me.Hide
    Else
        MsgBox "用户名或密码错误！"
    End If

End Sub

Function ClickOk()
    Call cmdOK_Click
End Function

Private Sub Command1_Click()
    AddUser.Show 1

    DoEvents
    AddUser.refre
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseTable

    End

End Sub

Function clearAutoSign() 'ok at 11-10-28
  eUser.Db.ExecNonQuery "Update Users Set isUsed = False"
End Function

Function SaveUserSetting(ByVal user As String) 'ok at 11-10-28
  Dim sql As String
  Dim bAutologin As Boolean
  Dim bRememberPass As Boolean

  sql = "Update `Users` Set `autologin` = ?, `rememberPass` = ?, `isUsed` = ? Where `uName` = ?"
  bAutologin = IIf(autologin.Value > 0, True, False)
  bRememberPass = IIf(remember.Value > 0, True, False)

  eUser.Db.ExecParamNonQuery sql, bAutologin, bRememberPass, True, user
End Function

Function ExistUser(ByVal user As String, ByVal pass As String) As Boolean
  Dim sql As String
  sql = "Select Count(*) From Users Where uName = ? And uPass = ?"
  If eUser.Db.ExecParamQueryScalar(sql, user, ReWind(pass)) > 0 Then
    ExistUser = True
  Else
    ExistUser = False
  End If
End Function

Function ExistUserM(ByVal user As String) As Boolean
  ExistUserM = eUser.UserExist(user)
End Function

Private Sub remember_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    If remember.Value = 1 Then
        autologin.Value = 0
    End If

End Sub

Private Sub user_Click() 'ok at 11-10-29

    If ExistUserM(user.Text) = True Then
        Call setuserState(user.Text)
    Else
        MsgBox "系统错误！", , "Sorry"
    End If

End Sub

Function setuserState(ByVal user As String) 'ok at 11-10-29
  If eUser.UserExist(user) Then
    Set res = eUser.Where("`uName` = ?", user)
    If res.fields("autologin") = True Then
      autologin.Value = 1
      remember.Value = 1
      pass.Text = res.fields("uPass")
      Exit Function
    End If
    If res.fields("rememberPass") = True Then
      remember.Value = 1
      pass.Text = res.fields("uPass")
    Else
      remember.Value = 0
      pass.Text = ""
    End If
  Else
    MsgBox "未知错误"
  End If
End Function

