VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBDocument"
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
Private m_UserName As String
Private m_IsChain As Boolean


Private Sub Class_Initialize()
  m_IsChain = False
  m_TableName = "Documents"
  m_TableFields = "Topic:text,AddTime:date,Remark:text," & _
  "Source:text,Class:string,IdNum:string," & _
  "auser:string,Content:longbinary,txtContent:text"
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

Public Function OfUser(ByVal UserName As String) As DBDocument
  m_UserName = UserName
  Set OfUser = Me
End Function

Public Function All() As ADODB.Recordset
  Set All = m_DBM.Db.ExecParamQuery("select * from Documents where auser = ?", m_UserName)
End Function

Public Function FindByID(ByVal IdNum As String) As ADODB.Recordset
  Set FindByID = m_DBM.Db.ExecParamQuery("select Topic,auser,AddTime,Class,IdNum,Content,Remark from Documents where IdNum = ? and auser = ?", IdNum, m_UserName)
End Function

Public Function FindByClassAndName(ByVal ClassName As String, ByVal UserName As String) As ADODB.Recordset
  Set FindByClassAndName = m_DBM.Db.ExecParamQuery("select * from Documents where Class = ? and auser = ?", ClassName, UserName)
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

Public Sub TransferToClass(ByVal ClassName As String, ByVal IdNum As String)
  m_DBM.Db.ExecParamNonQuery "Update Documents Set `Class` = ? Where `IdNum` = ?", ClassName, IdNum
End Sub

Public Sub Delete(ByVal IdNum As String)
  m_DBM.Db.ExecParamNonQuery "Delete From `Documents` Where `IdNum` = ?;", IdNum
End Sub

Public Sub Create(ByVal Topic As String, _
                  ByRef content() As Byte, _
                  ByVal Source As String, _
                  ByVal ClassName As String, _
                  ByVal IdNum As String, _
                  ByVal UserName As String, _
                  ByVal txtContent As String)
  Dim sql As String
  Dim addtime As String, Remark As String
  addtime = Format(Date, "yyyy-mm-dd")
  Remark = "NULL"
  
  sql = "INSERT INTO `Documents` (`Topic`, `AddTime`, `Remark`, `Source`," & _
        "`Class`, `IdNum`, `auser`, `Content`, `txtContent`) VALUES (" & _
        "?, ?, ?, ?," & _
        "?, ?, ?, ?, ?);"
  m_DBM.Db.ExecParamNonQuery sql, Topic, addtime, Remark, Source, _
                              ClassName, IdNum, UserName, content, txtContent
End Sub
