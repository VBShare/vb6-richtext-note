Attribute VB_Name = "MainModule"
Public Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As String, ByVal bEnable As Long, ByVal bCurrentThread As Long, ByRef bEnabled As Long) As Long
Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Public Declare Function SkinH_AttachEx _
               Lib "SkinH_VB6.dll" (ByVal lpSkinFile As String, _
                                    ByVal lpPasswd As String) As Long

Public LoginOk  As Boolean

Public conn     As ADODB.Connection

Public res      As ADODB.Recordset

Public nowLogin As String

Public art      As ADODB.Recordset

Public rartxt   As ADODB.Recordset

Public isDBon   As Boolean

Public adh As New AdodbHelper

Public Const SE_DEBUG_PRIVILEGE As Long = 20
'DataBase Entitys
Public documents As DBModel
Public class_ofs As DBModel
Public relates As DBModel
Public users As DBModel

Public eDocument As New DBDocument
Public eClassOf As New DBClassOf
Public eRelate As New DBRelate
Public eUser As New DBUser

Function OpenTable(ByVal txtPath As String) '【功能：建立数据库连接；状态：完成】
  Set conn = New ADODB.Connection
  conn.CursorLocation = adUseClient
  conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtPath & ";"
End Function

Function CloseTable() '【功能：关闭数据库连接；状态：完成】
  conn.Close
End Function

Public Sub CreateDb(ByVal DbPath As String)
  Dim dbc As New DbCreateHelper
  dbc.SetDbFile DbPath
  dbc.InitDbFromModels documents, class_ofs, relates, users
End Sub

Sub Main()
    '检查数据库文件是否存在
    'SkinH_AttachEx App.Path & "\QQ2011.she", "" '最后编译时设置为可见
    Dim DbPath As String
    Dim i As Integer
    DbPath = Replace(App.Path & "\documents.mdb", "\\", "\")
    Call RtlAdjustPrivilege(SE_DEBUG_PRIVILEGE, 0, 0, 1)
    Set documents = New DBDocument
    Set users = New DBUser
    Set class_ofs = New DBClassOf
    Set relates = New DBRelate

    If Dir(DbPath) = "" Then 'file does not exist
      CreateDb DbPath
    End If

    Call OpenTable(App.Path & "\documents.mdb")
    Load SearchMe
    Load Article
    Load NewArticle
    NewArticle.Timer1.Enabled = False
    Article.Hide
    Load login
    Load ArticleEdit
    Load AKind
    login.Hide
    Load AddUser
    Set res = New ADODB.Recordset
    res.Open "Users", conn, adOpenStatic, adLockOptimistic

    If res.RecordCount = 0 Then
        res.Close
        login.Show

        Exit Sub

    End If

    Do While Not res.EOF = True

        If CNull(res.fields("uName")) = "" Or CNull(res.fields("uPass")) = "" Then
            res.Delete adAffectCurrent
        Else
            login.user.AddItem res.fields("uName")

            If res.fields("autologin") = True Then
                login.autologin.Value = 1
            Else
                login.autologin.Value = 0
            End If

            If res.fields("rememberPass") = True Then
                login.autologin.Value = 1
            Else
                login.autologin.Value = 0
            End If

            If res.fields("autologin") = True And res.fields("isUsed") = True Then
                login.Show
                login.user.Text = res.fields("uName")
                login.pass.Text = res.fields("uPass")
                res.Close
                login.ClickOk

                Exit Sub

            Else
            End If
        End If

        res.MoveNext
    Loop

    login.autologin.Value = 0 And login.remember.Value = 0
    res.Close
    login.Show
End Sub

Function ReWindDoc(ByVal inPutX As String) 'ok at 11-10-28
    'and add limits
    inPutX = Replace(inPutX, "'", "''")
    ReWindDoc = inPutX
End Function

Function ReWind(ByVal inPutX As String) 'ok at 11-10-28
    'and add limits
    inPutX = Replace(inPutX, "'", "''")

    If Len(inPutX) > 30 Then
        ReWind = Mid(inPutX, 1, 30)
    Else
        ReWind = inPutX
    End If

End Function

Function CNull(ByVal sTxt As Variant) As String 'ok at 11-10-08

    If IsNull(sTxt) = True Then
        CNull = ""
    Else
        CNull = sTxt
    End If

End Function

Function isInClass(ByVal ClassName As String) As Boolean '检查某个分类是否存在ok at 11-10-29
    Set res = New ADODB.Recordset
    res.Open "select * from ClassOf where className='" & ClassName & "'", conn, 3, 3

    If res.RecordCount = 0 Then
        res.Close
        isInClass = False

        Exit Function

    End If

    res.Close
    isInClass = True
End Function

Function OutputFileS(ByVal sId As Long, ByVal sFile As String) 'ok at 11-10-29

    Dim sTemp() As Byte

    sTemp = LoadResData(sId, "CUSTOM")
    Open sFile For Binary As #1
    Put #1, , sTemp
    Close #1
End Function

Function Service(ByVal cmd As String) As Boolean 'ok at 11-10-29

    If cmd = "" Then
        Service = False

        Exit Function

    End If

    Select Case cmd

        Case "RefreshClass"
            Call AddUser.refre
            Call Article.RefreshClass
            Call ArticleEdit.RefreshClass
            Call NewArticle.RefreshClass
            Service = True

        Case "RefreshUser"
            Call loadUser

        Case "RefreshDocuments"
            Call Article.LoadArticle(nowLogin)
            Call ArticleEdit.LoadArticle(nowLogin)

        Case "InputArticle"
            Call Article.ClsArticle

        Case "InputArticleEdit"
            Call ArticleEdit.ClsArticle

        Case "InputNewArticle"
            Call NewArticle.ClsArticle

        Case "SectionKey"
            Call AKind.RefreshKind
    End Select

End Function

Function loadUser() 'ok at 11-10-29
  Set res = eUser.All
  If res.RecordCount = 0 Then
      adh.ReleaseRecordset res
      Exit Function
  End If

  login.user.Clear

  Do While Not res.EOF = True
    If CNull(res.fields("uName")) = "" Or CNull(res.fields("uPass")) = "" Then
      res.Delete adAffectCurrent
    Else
      login.user.AddItem res.fields("uName")

      If res.fields("autologin") = True Then
          login.autologin.Value = 1
      Else
          login.autologin.Value = 0
      End If

      If res.fields("rememberPass") = True Then
          login.autologin.Value = 1
      Else
          login.autologin.Value = 0
      End If
    End If

    res.MoveNext
  Loop
  
  adh.ReleaseRecordset res
End Function

Function AddClass(ByVal ClassName As String, ByVal UserName) 'ok at 11-12-29
  eClassOf.Create ClassName, UserName
End Function

Function AutoSection(ByVal Str As String) 'ok at 11-11-06
  Dim i As Integer

  Set res = adh.ExecQuery("select * from Relate")

  If res.RecordCount = 0 Then
      adh.ReleaseRecordset res
      Exit Function
  End If

  Do While Not res.EOF = True
    If InStr(1, LCase(Str), LCase(res.fields("keyword"))) > 0 Then
      Exit Do
    End If

    res.MoveNext
  Loop

  With NewArticle

      For i = 0 To .section.ListCount - 1

          If InStr(1, res.fields("section"), .section.List(i)) > 0 Then
              .section.ListIndex = i

              Exit For

          End If

      Next i

  End With

  adh.ReleaseRecordset res
End Function

Function OpenURL(ByVal url As String)

    Dim lngReturn As Long

    lngReturn = ShellExecute(0, "open", url, "", "", 0)
End Function
