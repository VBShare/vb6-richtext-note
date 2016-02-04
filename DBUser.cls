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
Implements DbModel
'Private
Private m_TableName As String
Private m_TableFields As String
Private m_DBM As New DbModel
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


