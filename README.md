# vb6-richtext-note
a rich text dairy

This is a project where AdodbHelper is applied.

AdodbHelper is an VB6 Class Designed to avoid the problem of manage database connection.

Hope that well help.

# Demo Of How To Create DataBase From Code
# 如何使用代码创建数据库的一个演示
我们需要用到的是一个基类DBModel
现在，我们需要建立一个Documents表，我们就先建立一个类：DBDocument

代码如下：
```vb
Option Explicit
Implements DBModel
'Private
Private m_TableName As String
Private m_TableFields As String
Private m_DBH As New AdodbHelper
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
```
在类中，可以看到
```vb
Private Sub Class_Initialize()
  m_IsChain = False
  m_TableName = "Documents"
  m_TableFields = "Topic:text,AddTime:date,Remark:text," & _
  "Source:text,Class:string,IdNum:string," & _
  "auser:string,Content:longbinary,txtContent:text"
End Sub
```
这里的m_TableName的值就是表名

这里的m_TableFields的值就是字段列表

`title:string`表示一个string类型的字段，名字叫title。
每种类型对应的意义，如下表
| 类型      |  说明    |
| :-------- | :--------|
|string|VarChar(255)（adVarWChar）|
|text|长字符串，2亿字符（adLongVarWChar）|
|date|日期（adDate）|
|currency|金额，高精度、长度长（adCurrency）|
|boolean|Bit，逻辑值（adBoolean）|
|double|双精度数（adDouble）|
|integer|整数，2字节（adInteger）|
|guid|GUID（adGUID）|
|single|单精度数（adSingle）|
|longbinary|字节数组（adLongVarBinary）|
|byte|Byte(255)（adUnsignedTinyInt）|
|short|Byte字节（adSmallInt）|

有了类，如何创建数据库呢？我们有一个DbCreateHelper用来创建数据库，如下
```vb
Dim dbc As New DbCreateHelper               '一个DbCreateHelper实例
Dim mDocument As DBModel                    '定义使用DBModel接口的mDocument
Set mDocument = New DBDocument              '建立DBDocument对象，赋给mDocument
dbc.SetDbFile App.Path & "\document.mdb"    '设定数据库文件的输出路径
dbc.InitDbFromModels mDocument              '创建数据库
```
