VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements DBModel
'Private
Private m_TableName As String
Private m_TableFields As String
Private m_DBM As New DBModel
Private Sub Class_Initialize()
  m_TableName = "Users"
  m_TableFields = "uName:string,uPass:string,autologin:boolean," & _
  "rememberPass:boolean,isUsed:boolean"
End Sub

Private Property Get DbModel_Db() As AdodbHelper
  Set DbModel_Db = m_DBM.Db
End Property

Private Property Get DbModel_TableFields() As String
  DbModel_TableFields = m_TableFields
End Property

Private Property Get DbModel_TableName() As String
  DbModel_TableName = m_TableName
End Property

Public Function UserExist(ByVal UserName As String) As Boolean
  If m_DBM.Db.ExecParamQueryScalar("select count(*) from Users where uName= ? ", UserName) > 0 Then
    UserExist = True
  Else
    UserExist = False
  End If
End Function

Public Sub Create(ByVal UserName As String, ByVal PassWord As String)
  m_DBM.Db.ExecParamNonQuery "Insert Into Users(uName, uPass, autologin, rememberPass, isUsed) values(?, ?, ?, ?, ?)", UserName, PassWord, False, False, False
End Sub

Public Function All() As ADODB.Recordset
  Set All = m_DBM.Db.ExecQuery("select * from Users")
End Function

Private Function DBModel_Where(ByVal Conditions As String, ParamArray Params() As Variant) As ADODB.Recordset
  Dim sql As String
  If Len(Conditions) = 0 Then
    sql = "select * from " & m_TableName
  Else
    sql = "select * from " & m_TableName & " where " & Conditions
  End If
  Set DBModel_Where = m_DBM.Db.ExecParamQuery(sql, Params)
End Function
