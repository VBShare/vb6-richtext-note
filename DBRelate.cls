VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBRelate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements DBModel
'Private
Private m_TableName As String
Private m_TableFields As String
Private m_DBH As New AdodbHelper
Private Sub Class_Initialize()
  m_TableName = "Relate"
  m_TableFields = "section:text,keyword:text"
End Sub

Private Property Get DbModel_Db() As AdodbHelper
  Set DbModel_Db = m_DBH
End Property
Public Property Get Db() As AdodbHelper
  Set Db = DbModel_Db
End Property

Private Sub DBModel_InitConn(ByVal dbFilePath As String)
  m_DBH.SetConnToFile dbFilePath
End Sub
Public Sub InitConn(ByVal dbFilePath As String)
  DBModel_InitConn dbFilePath
End Sub

Private Property Get DbModel_TableFields() As String
  DbModel_TableFields = m_TableFields
End Property
Public Property Get TableFields() As String
  TableFields = DbModel_TableFields
End Property

Private Property Get DbModel_TableName() As String
  DbModel_TableName = m_TableName
End Property
Public Property Get TableName() As String
  TableName = DbModel_TableName
End Property


Public Sub Create(ByVal Keyword As String, ByVal section As String)
  Dim sql As String
  sql = "Insert Into `Relate`(`keyword`, `section`) values (?, ?)"
  m_DBM.Db.ExecParamNonQuery sql, Keyword, section
End Sub

Public Function KeywordExist(ByVal Keyword As String) As Boolean
  Dim sql As String
  sql = "select count(*) from `Relate` where `keyword` = ?"
  If m_DBM.Db.ExecParamQueryScalar(sql, Keyword) > 0 Then
    KeywordExist = True
  Else
    KeywordExist = False
  End If
End Function

Public Sub RemoveRelate(ByVal section As String, ByVal Keyword As String)
  Dim sql As String
  sql = "Delete From `Relate` Where `keyword` = ? And `section` = ?"
  m_DBM.Db.ExecParamNonQuery sql, Keyword, section
End Sub

Public Sub RemoveRelateOfClass(ByVal section As String)
  Dim sql As String
  sql = "Delete From `Relate` Where `section` = ?"
  m_DBM.Db.ExecParamNonQuery sql, section
End Sub

Public Function OfClass(ByVal section As String) As ADODB.Recordset
  Dim sql As String
  sql = "Select * From `Relate` Where `section` = ?"
  Set OfClass = m_DBM.Db.ExecParamQuery(sql, section)
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
Public Function Where(ByVal Conditions As String, ParamArray Params() As Variant) As ADODB.Recordset
  Dim sql As String
  If Len(Conditions) = 0 Then
    sql = "select * from " & m_TableName
  Else
    sql = "select * from " & m_TableName & " where " & Conditions
  End If
  Set Where = m_DBM.Db.ExecParamQuery(sql, Params)
End Function
