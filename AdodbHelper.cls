VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdodbHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'����޸����ڣ�2015��1��15��22:16:43
'2015��1��15��22:19:23      ȫ�����¸�ʽ��

'��Ա���������Ӷ���
Private m_Conn       As ADODB.Connection

'��Ա������SQL�������
Private m_Command    As ADODB.Command

'��Ա�����������ַ���
Private m_ConnString As String

'��Ա���������ݿ��ļ�·��
Private m_FilePath   As String

'��Ա��������������
Private m_Me         As New AdodbHelper

'�涨��ÿ������ִ����ɺ󣬶��������m_Command�����Ҵ���һ���յĶ���
'�����ʼ�����ͷš�
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       Class_Initialize
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Initialize()
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       Class_Terminate
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Class_Terminate()
    Set m_Conn = Nothing
    Set m_Command = Nothing
End Sub

'�������ԡ�
'��ȡ�����ݿ������ַ���
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ConnectionString
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get ConnectionString() As String
    ConnectionString = m_ConnString
End Property

'���ã������ַ���
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ConnectionString
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :       vNewValue (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Let ConnectionString(ByVal vNewValue As String)
    m_ConnString = vNewValue
End Property

'���Ƿ�׼���ã��Ƿ����ִ��SQL
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       IsReady
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:56
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get IsReady() As Boolean
    IsReady = IIf(Len(ConnectionString) > 0, True, False)
End Property

'���෽����
'ִ�в�ѯ���������߼�¼��
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ExecQuery
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecQuery(ByVal SqlStr As String) As ADODB.Recordset

    Dim tempRes As New Recordset

    Set m_Command = New ADODB.Command
    Call OpenConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    Set tempRes = m_Command.Execute()
    tempRes.ActiveConnection = Nothing
    Call CloseConn
    Set ExecQuery = tempRes
    Set m_Command = Nothing
End Function

'ִ�в�������ѯ���������߼�¼��
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ExecParamQuery
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'                    Params() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecParamQuery(ByVal SqlStr As String, _
                               ParamArray Params()) As ADODB.Recordset

    Dim tempRes As New Recordset

    Dim i       As Long

    Set m_Command = New ADODB.Command
    '������
    Call OpenConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    m_Command.CommandType = adCmdText

    '���ò���
    With m_Command

        For Each param In Params

            Dim Para As ADODB.Parameter

            Set Para = .CreateParameter(CStr(i), GetVarType(param), adParamInput, LenB(param))
            Para.Value = param
            .Parameters.Append Para
        Next

    End With

    '��ȡִ�к��¼��
    Set tempRes = m_Command.Execute()
    '�����ݿ������ѹ�
    tempRes.ActiveConnection = Nothing
    '�ر����ݿ�����
    Call CloseConn
    '�������ݼ���������
    Set ExecParamQuery = tempRes
    '����������
    Set m_Command = Nothing
End Function

'ִ�в�ѯ������Ӱ������
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ExecNonQuery
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecNonQuery(ByVal SqlStr As String) As Long

    '����Ӱ����������
    Dim affectedRows As Long

    '������Command����
    Set m_Command = New ADODB.Command
    '������
    Call OpenConn
    '��Command�����ݿ�����
    m_Command.ActiveConnection = m_Conn
    '����SQL���
    m_Command.CommandText = SqlStr
    '����SQL����
    m_Command.CommandType = adCmdText
    '��ȡִ�к�Ӱ������
    m_Command.Execute affectedRows
    '�ر����ݿ�����
    Call CloseConn
    '����������
    Set m_Command = Nothing
    '����Ӱ������
    ExecNonQuery = affectedRows
End Function

'ִ�в�������ѯ������Ӱ������
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ExecParamNonQuery
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'                    Params() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecParamNonQuery(ByVal SqlStr As String, ParamArray Params()) As Long

    Dim i            As Long

    Dim affectedRows As Long

    Set m_Command = New ADODB.Command
    '������
    Call OpenConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    m_Command.CommandType = adCmdText

    '���ò���
    With m_Command

        For Each param In Params

            Dim Para As ADODB.Parameter

            Set Para = .CreateParameter(CStr(i), GetVarType(param), adParamInput, LenB(param))
            Para.Value = param
            .Parameters.Append Para
        Next

    End With

    '��ȡִ�к��¼��
    m_Command.Execute affectedRows
    '�ر����ݿ�����
    Call CloseConn
    '����������
    Set m_Command = Nothing
    '����Ӱ������
    ExecParamNonQuery = affectedRows
End Function

'�����ļ�·�����������ַ���
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       SetConnToFile
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       FilePath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub SetConnToFile(ByVal FilePath As String)
    m_ConnString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
End Sub

'�ͷ����߼�¼��
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ReleaseRecordset
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       dbRes (ADODB.Recordset)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub ReleaseRecordset(ByRef dbRes As ADODB.Recordset)
    Set dbRes = Nothing
End Sub

'���ƶ������ݴ�һ�����ݿ��ļ����뵽��һ�����ݿ��ļ���
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       TransData
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SourceDb (String)
'                    DestDb (String)
'                    tableNames() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub TransData(ByVal SourceDb As String, _
                     ByVal DestDb As String, _
                     ParamArray tableNames())

    Dim tableName As Variant

    m_Me.SetConnToFile SourceDb

    If UBound(tableNames) <= 0 Then

        Exit Sub

    End If

    For Each tableName In tableNames

        m_Me.ExecQuery "select * into [" & DestDb & "]." & tableName & " from " & tableName
    Next

End Sub

'�Ӵ����ݿ��ļ�������һ���򿪵����ݿ�����
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       DbConnFromFile
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       FilePath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function DbConnFromFile(ByVal FilePath As String) As ADODB.Connection

    Dim tmpConn As New ADODB.Connection

    tmpConn.CursorLocation = adUseClient
    tmpConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
    Set DbConnFromFile = tmpConn
End Function

'ִ�в�������ѯ���������У�����ֵ
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ExecParamQueryScalar
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'                    Params() (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecParamQueryScalar(ByVal SqlStr As String, _
                                     ParamArray Params()) As Variant

    Dim varResult As Variant

    Dim tmpRes    As Recordset

    Set tmpRes = ExecParamQuery(SqlStr, Params(0))

    If tmpRes.RecordCount <= 0 Then
        Set ExecParamQueryScalar = Nothing
    Else
        ExecParamQueryScalar = tmpRes.Fields(0).Value
    End If

    ReleaseRecordset tmpRes
End Function

'ִ�в�ѯ���������У�����ֵ
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       ExecQueryScalar
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       SqlStr (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ExecQueryScalar(ByVal SqlStr As String) As Variant

    Dim varResult As Variant

    Dim tmpRes    As Recordset

    Set tmpRes = ExecQuery(SqlStr)

    If tmpRes.RecordCount <= 0 Then
        Set ExecQueryScalar = Nothing
    Else
        ExecQueryScalar = tmpRes.Fields(0).Value
    End If

    ReleaseRecordset tmpRes
End Function

'���ڲ�������
'�����ݿ�����
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       OpenConn
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub OpenConn()
    Set m_Conn = New ADODB.Connection
    m_Conn.CursorLocation = adUseClient
    m_Conn.Open ConnectionString
End Sub

'�ر����ݿ�����
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       CloseConn
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CloseConn()
    m_Conn.Close
    Set m_Conn = Nothing
End Sub

'����VB�������ͷ������ݿ����������ö��ֵ
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ��⹹����Ŀ
' Procedure  :       GetVarType
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       MYSOFT-SUNR01
' Date-Time  :       1-18-2015-18:19:55
'
' Parameters :       Value (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function GetVarType(ByRef Value As Variant) As DataTypeEnum

    Dim k As New ADODB.Command

    Select Case VarType(Value)

        Case VbVarType.vbString
            GetVarType = DataTypeEnum.adVarChar

        Case VbVarType.vbInteger
            GetVarType = DataTypeEnum.adSmallInt

        Case VbVarType.vbBoolean
            GetVarType = DataTypeEnum.adBoolean

        Case VbVarType.vbCurrency
            GetVarType = DataTypeEnum.adCurrency

        Case VbVarType.vbDate
            GetVarType = DataTypeEnum.adDate

        Case 8209
            GetVarType = DataTypeEnum.adLongVarBinary

        Case Else
            GetVarType = DataTypeEnum.adVariant
    End Select

End Function
