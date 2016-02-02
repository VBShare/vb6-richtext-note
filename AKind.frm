VERSION 5.00
Begin VB.Form AKind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ؼ��������"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "ɾ��"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ɾ��(&D)"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���(&A)"
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
      Caption         =   "��Ӧ�ؼ���"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����"
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

    If List1.ListIndex < 0 Then MsgBox "��ѡ�����": Exit Sub
    s = InputBox("������ؼ���", "�ؼ�������")

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

Function AddRelate(ByVal keyword As String, sections As String) 'ok at 11-11-06
    '
    Set res = New ADODB.Recordset
    res.Open "select * from Relate where keyword='" & keyword & "'", conn, 3, 3

    If res.RecordCount > 0 Then MsgBox "��ͬ�ؼ��ִ��ڣ������ظ����", vbCritical, "����": res.Close: Exit Function
    res.Close

    'begin add
    Set res = New ADODB.Recordset
    res.Open "Relate", conn, 3, 3

    With res
        .AddNew
        .Fields("keyword") = keyword
        .Fields("section") = sections
        .Update
    End With

    res.Close
End Function

Function DelRelate(ByVal keyword As String, sections As String) 'ok at 11-11-06
    Set res = New ADODB.Recordset
    res.Open "delete * from Relate where section='" & sections & "' and keyword='" & keyword & "'", conn, 3, 3
End Function

Function DelClass(ByVal clsname As String) 'ok at 11-12-07
    Set res = New ADODB.Recordset
    res.Open "delete * from ClassOf where className='" & clsname & "'", conn, 3, 3
    Set res = New ADODB.Recordset
    res.Open "delete * from Relate where section='" & clsname & "'", conn, 3, 3
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
    Set res = New ADODB.Recordset
    res.Open "select * from Relate where [section]='" & List1.List(List1.ListIndex) & "'", conn, adOpenStatic, adLockOptimistic

    If res.RecordCount = 0 Then res.Close: Exit Sub

    Do While Not res.EOF = True
        List2.AddItem res.Fields("keyword")
        res.MoveNext
    Loop

    res.Close
End Sub

Function RefreshKind() 'ok at 11-11-06
    List1.Clear
    List2.Clear
    Set res = New ADODB.Recordset
    res.Open "ClassOf", conn, 3, 3

    If res.RecordCount = 0 Then res.Close: Exit Function

    Do While Not res.EOF = True
        List1.AddItem res.Fields("className")
        res.MoveNext
    Loop

    res.Close
End Function
