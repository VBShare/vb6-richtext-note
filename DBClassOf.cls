VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClassOf"
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
  m_TableName = "ClassOf"
  m_TableFields = "className:string,userName:string"
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

Public Sub Create(ByVal ClassName As String, ByVal UserName As String)
  Dim sql As String
  sql = "INSERT INTO `ClassOf`(`className`, `userName`) Values(?, ?)"
  m_DBH.ExecParamNonQuery sql, ClassName, UserName
End Sub

Public Sub RemoveClass(ByVal ClassName As String)
  m_DBH.ExecParamNonQuery "Delete From `ClassOf` Where `className` = ?", ClassName
End Sub

Public Function All() As ADODB.Recordset
  Set All = m_DBH.ExecQuery("Select * From `ClassOf`")
End Function

Private Function DBModel_Where(ByVal Conditions As String, ParamArray Params() As Variant) As ADODB.Recordset
  Dim sql As String
  If Len(Conditions) = 0 Then
    sql = "select * from " & m_TableName
  Else
    sql = "select * from " & m_TableName & " where " & Conditions
  End If
  Set DBModel_Where = m_DBH.ExecParamQuery(sql, Params)
End Function

Public Function Where(ByVal Conditions As String, ParamArray Params() As Variant) As ADODB.Recordset
  Dim sql As String
  If Len(Conditions) = 0 Then
    sql = "select * from " & m_TableName
  Else
    sql = "select * from " & m_TableName & " where " & Conditions
  End If
  Set Where = m_DBH.ExecParamQuery(sql, Params)
End Function
