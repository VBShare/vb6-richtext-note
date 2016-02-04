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
Implements DbModel
'Private
Private m_TableName As String
Private m_TableFields As String
Private m_DBM As New DbModel
Private m_UserName As String


Private Sub Class_Initialize()
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
