VERSION 5.00
Begin VB.Form AddUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�¹���Ա"
   ClientHeight    =   2565
   ClientLeft      =   7605
   ClientTop       =   4080
   ClientWidth     =   3195
   Icon            =   "AddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3195
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton btnNewUser 
      Caption         =   "���"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox UserName 
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox RePassWord 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox PassWord 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "�ظ�����"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "�û���"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "AddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnNewUser_Click() 'add new ok at 11-10-29

    If UserName.Text = "" Then MsgBox "�������û���": Exit Sub
    If PassWord.Text = "" Then MsgBox "����������": Exit Sub
    If RePassWord.Text = "" Then MsgBox "���ٴ���������": Exit Sub
    If PassWord.Text <> RePassWord.Text Then MsgBox "������������벻ͬ��": Exit Sub

    '���
    If eUser.UserExist(UserName.Text) Then
      MsgBox "��ͬ�û�������"
      Exit Sub
    Else
      GoTo addit
    End If

addit:
  eUser.Create ReWind(UserName.Text), ReWind(PassWord.Text)
  MsgBox "��ӳɹ���"
  Call Service("RefreshUser")
  Me.Hide
End Sub

Private Sub Command2_Click() 'cancel ok at 11-10-28
  AddUser.Hide
End Sub

Private Sub PassWord_Change() 'ok at 11-10-28
  Call Limit(PassWord)
End Sub

Private Sub RePassWord_Change() 'ok at 11-10-28
  Call Limit(RePassWord)
End Sub

Private Sub UserName_Change() 'ok at 11-10-28
  Call Limit(UserName)
End Sub

Function Limit(ByRef obj As Object) 'ok at 11-10-28
  If Len(obj.Text) > 29 Then
    obj.Text = Mid(obj.Text, 1, 30)
    obj.SelStart = Len(obj.Text)
  End If
End Function

Function refre()
  UserName.Text = ""
  PassWord.Text = ""
  RePassWord.Text = ""
End Function
