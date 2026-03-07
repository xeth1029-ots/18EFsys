Public Class cls_Exam

    Public Shared Function Get_ExamTypePName(ByVal PETID As String, ByVal DistID As String, ByVal tConn As SqlConnection) As String
        Dim Rst As String = ""
        'PARENT IS NULL 要做為父層 不可有父層 (系統限制2層)
        Dim sql As String = "SELECT * FROM ID_ExamType WHERE PARENT IS NULL AND DistID ='" & DistID & "'"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, tConn)
        If TIMS.dtHaveDATA(dt) Then
            If dt.Select("ETID=" & PETID).Length > 0 Then
                Rst = Convert.ToString(dt.Select("ETID=" & PETID)(0)("Name"))
            End If
        End If
        Return Rst
    End Function

    Public Shared Function Get_ExamTypeParent(ByVal obj As ListControl, ByVal DistID As String, Optional ByVal sType As Integer = 0) As ListControl
        'sType 0'上層類別 1'請選擇 'Avail=1 啟用
        Using conn As SqlConnection = DbAccess.GetConnection()
            Dim pms1 As New Hashtable From {{"DistID", DistID}}
            Dim sql As String = "SELECT * FROM ID_ExamType WHERE PARENT IS NULL AND Avail=1 AND DistID=@DistID"
            Dim dt As DataTable = DbAccess.GetDataTable(sql, conn, pms1)
            With obj
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "ETID"
                .DataBind()
                Select Case sType
                    Case 0
                        If TypeOf obj Is DropDownList Then .Items.Insert(0, New ListItem("==上層類別==", ""))
                    Case 1
                        If TypeOf obj Is DropDownList Then .Items.Insert(0, New ListItem("==請選擇==", ""))
                End Select
            End With
        End Using
        Return obj
    End Function

    '取出鍵詞-子類別代碼
    Public Shared Function Get_ExamType(ByVal obj As ListControl, ByVal DistID As String, ByVal ParentETID As Integer) As ListControl
        Dim pms1 As New Hashtable From {{"ParentETID", ParentETID}}
        Dim sql As String = ""
        sql &= " select c.etid , c.name as Name" & vbCrLf
        sql &= " FROM ID_ExamType c" & vbCrLf
        sql &= " where C.avail=1 and c.Parent =@ParentETID " & vbCrLf
        '系統管理者可查全部
        If DistID <> "000" Then sql &= " and C.DistID = '" & DistID & "'" & vbCrLf
        sql &= " ORDER BY 2" & vbCrLf
        Using conn As SqlConnection = DbAccess.GetConnection()
            Dim dt As DataTable = DbAccess.GetDataTable(sql, conn, pms1)
            With obj
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "ETID"
                .DataBind()
                If TypeOf obj Is DropDownList Then
                    .Items.Insert(0, New ListItem("不區分", Convert.ToString(ParentETID)))
                End If
            End With
        End Using

        Return obj
    End Function

    Public Shared Function Get_qtypeName(ByVal qtypeID As String) As String
        Dim qtypeName As String = ""
        Select Case qtypeID
            Case "1"
                qtypeName = "是非題"
            Case "2"
                qtypeName = "選擇題"
            Case "3"
                qtypeName = "複選題"
            Case "4"
                qtypeName = "問答題"
        End Select
        Return qtypeName
    End Function

    Public Shared Sub SetQTypeName(ByVal obj As ListControl, ByVal IsOnline As String)
        '設定題目類型
        With obj
            .Items.Clear()
            .Items.Insert(0, New ListItem("==請選擇==", "0"))
            .Items.Insert(1, New ListItem("是非題", "1"))
            .Items.Insert(2, New ListItem("選擇題", "2"))
            .Items.Insert(3, New ListItem("複選題", "3"))
            If IsOnline = "N" Then
                'SetQType(ddl_qtype, dr("isonline"))
                '非線上考試，可考問答題
                .Items.Insert(4, New ListItem("問答題", "4"))
            End If
        End With
    End Sub

#Region "NO USE"
    '取出鍵詞-上層類別代碼
    'Public Shared Function Get_ExamTypeName(ByVal ETIDval As String) As String
    '    Dim ExamTypeName As String
    '    Dim dr As DataRow
    '    Dim Sql As String
    '    Sql = "select name from id_examtype where etid=" & ETIDval
    '    dr = DbAccess.GetOneRow(Sql)
    '    If Not dr Is Nothing Then
    '        ExamTypeName = dr("name")
    '    Else
    '        ExamTypeName = "不區分"
    '    End If
    '    Return ExamTypeName
    'End Function
    '取出鍵詞-上層類別代碼
#End Region

End Class
