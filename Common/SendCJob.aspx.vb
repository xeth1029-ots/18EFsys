Partial Class SendCJob
    Inherits AuthBasePage

    'SHARE_CJOB cjob2016
    Dim Rqfield As String = "" 'Request("field") '不定(欄位)
    Dim RqcjobValue As String = "" 'Request("cjobValue")
    'Const cst_2016 As String="2016" '啟動2016
    'Dim sCjob2016 As String="" 'sCjob2016=TIMS.Utl_GetConfigSet("cjob2016")
    Dim str_SHARECJOB_YEAR As String = ""

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        '檢查Session是否存在 End

#Region "(No Use)"

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then
        '    '若SESSION 異常
        '    Call TIMS.sUtl_404NOTFOUND(Me, objconn)
        '    Exit Sub
        'End If

#End Region

        '啟用2016年通俗職類-'啟動2019-命名2016
        Dim flag_Cjob2016 As Boolean = TIMS.Get_sCjob2016_USE(Me)
        'Dim str_SHARECJOB_YEAR As String=""
        str_SHARECJOB_YEAR = ""
        If flag_Cjob2016 Then str_SHARECJOB_YEAR = TIMS.cst_SHARE_CJOB_2016 '啟動2019-命名2016

        Dim sql As String = ""
        tbCJOB1.Visible = False
        tbCJOB2.Visible = False
        Select Case str_SHARECJOB_YEAR
            Case TIMS.cst_SHARE_CJOB_2016 '啟動2019-命名2016
                tbCJOB2.Visible = True
            Case Else
                '第1代通俗職類(OLD)
                tbCJOB1.Visible = True
        End Select

        Rqfield = TIMS.ClearSQM(Request("field")) '不定(欄位)
        RqcjobValue = TIMS.ClearSQM(Request("cjobValue")) '有接收值
        If TIMS.CheckInput(Rqfield) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
            Exit Sub
        End If
        If TIMS.CheckInput(RqcjobValue) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
            Exit Sub
        End If
        fieldname.Value = Rqfield

        If Not IsPostBack Then Call CCreate1()
    End Sub

    Sub CCreate1()
        'Const cst_2016 As String="2016"
        'Dim sCjob2016 As String=TIMS.Utl_GetConfigSet("cjob2016")
        Rqfield = TIMS.ClearSQM(Request("field")) '不定(欄位)
        RqcjobValue = TIMS.ClearSQM(Request("cjobValue")) '有接收值
        If TIMS.CheckInput(Rqfield) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
            Exit Sub
        End If
        If TIMS.CheckInput(RqcjobValue) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
            Exit Sub
        End If
        fieldname.Value = Rqfield

        '啟用2016年通俗職類
        'Dim flag_Cjob2016 As Boolean=TIMS.Get_sCjob2016_USE(Me)
        'Dim str_SHARECJOB_YEAR As String=""
        'If flag_Cjob2016 Then str_SHARECJOB_YEAR=TIMS.cst_SHARE_CJOB_2016

        Select Case str_SHARECJOB_YEAR
            Case TIMS.cst_SHARE_CJOB_2016 '啟動2019-命名2016
                'tbCJOB2.Visible=True
                ddlCJOB16A2.Items.Clear()
                ddlCJOB16A3.Items.Clear()
                ddlCJOB16A1 = GetCJobType(ddlCJOB16A1)
                'ddlCJOB16A1=GetCJobUNKEYNO(ddlCJOB16A1, CJOB_TYPE.SelectedValue, "")
                If RqcjobValue <> "" Then
                    Dim dr1 As DataRow = GetCJobTypeVal(RqcjobValue)
                    If dr1 Is Nothing Then Exit Sub
                    Dim vsCJOB_TYPE As String = Convert.ToString(dr1("cjob_type"))
                    Dim vsCJOB_NO As String = Convert.ToString(dr1("cjob_no"))
                    Dim vsCJOB_UNKEY As String = Convert.ToString(dr1("CJOB_UNKEY"))
                    Common.SetListItem(ddlCJOB16A1, vsCJOB_TYPE)
                    ddlCJOB16A2 = GetCJobUNKEYNO(ddlCJOB16A2, vsCJOB_TYPE, "2")
                    Common.SetListItem(ddlCJOB16A2, vsCJOB_NO)
                    ddlCJOB16A3 = GetCJobUNKEYNO(ddlCJOB16A3, vsCJOB_NO, "3")
                    Common.SetListItem(ddlCJOB16A3, vsCJOB_UNKEY)
                    'Dim CJobType1 As String=GetCJobTypeVal(RqcjobValue)
                    'Dim CJobType2 As String=GetCJobTypeVal(RqcjobValue)
                End If
            Case Else '第1代通俗職類
                CJOB_TYPE = GetCJobType(CJOB_TYPE)
                CJOB_NO = GetCJobUNKEYNO(CJOB_NO, CJOB_TYPE.SelectedValue, "")
                If RqcjobValue <> "" Then
                    Dim dr1 As DataRow = GetCJobTypeVal(RqcjobValue)
                    If dr1 Is Nothing Then Exit Sub
                    Dim vsCJOB_TYPE As String = Convert.ToString(dr1("cjob_type"))
                    If vsCJOB_TYPE <> "" Then Common.SetListItem(CJOB_TYPE, vsCJOB_TYPE)
                    CJOB_NO = GetCJobUNKEYNO(CJOB_NO, CJOB_TYPE.SelectedValue, "")
                    If RqcjobValue <> "" Then Common.SetListItem(CJOB_NO, RqcjobValue)
                End If
        End Select
    End Sub

#Region "common func"

    '設計 DropDownList 依 SHARE_CJOB 
    Function GetCJobType(ByVal obj As ListControl) As ListControl
        '啟動2019-命名2016'str_SHARECJOB_YEAR=TIMS.cst_SHARE_CJOB_2016 
        Dim sql As String = ""
        sql &= " SELECT CJOB_TYPE ,concat('[',CJOB_TYPE,']',CJOB_NAME) CJOB_NAME"
        sql &= " FROM SHARE_CJOB"
        sql &= " WHERE CJOB_NO IS NULL AND CYEARS='2019' AND CDEL IS NULL"
        sql &= " ORDER BY CONVERT(numeric, CJOB_TYPE)"
        '    sql &= " SELECT CJOB_TYPE ,CJOB_NAME FROM SHARE_CJOB"
        '    sql &= " WHERE CJOB_NO IS NULL AND CYEARS='2014'"
        '    sql &= " ORDER BY CONVERT(numeric, CJOB_TYPE)"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        With obj
            .DataSource = dt
            .DataTextField = "CJOB_NAME"
            .DataValueField = "CJOB_TYPE"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            'If TypeOf obj Is DropDownList Then .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        Return obj
    End Function

    '依 CJOB_TYPE 設計 DropDownList  
    Function GetCJobUNKEYNO(ByVal obj As ListControl, ByVal cjob_value As String, ByVal sType As String) As ListControl
        obj.Items.Clear()
        If cjob_value = "" Then Return obj

        'Case "" '2014 '第1代通俗職類
        'Dim CJOB_TYPE As String = cjob_value
        'sql &= " SELECT CJOB_UNKEY,CJOB_TYPE,'[' + CJOB_NO + ']' + CJOB_NAME JOBNAME" & vbCrLf
        'sql &= " FROM SHARE_CJOB" & vbCrLf
        'sql &= $" WHERE CJOB_NO IS NOT NULL AND CYEARS='2014' AND CJOB_TYPE='{CJOB_TYPE}'"
        'sql &= " ORDER BY CONVERT(numeric, CJOB_TYPE) ,CJOB_UNKEY" & vbCrLf
        Dim sql As String = ""
        Select Case sType
            Case "2" '2016 選第2層
                Dim vCJOB_TYPE As String = cjob_value
                'CJOB_NO->CJOB_UNKEY
                sql &= " SELECT CJOB_NO CJOB_UNKEY ,CJOB_TYPE ,'[' + CJOB_NO + ']' + CJOB_NAME jobName" & vbCrLf
                sql &= " FROM SHARE_CJOB" & vbCrLf
                sql &= $" WHERE CJOB_NO IS NOT NULL AND JOB_NO IS NULL AND CYEARS='2019' AND CJOB_TYPE='{vCJOB_TYPE}'"
                sql &= " ORDER BY CJOB_NO" & vbCrLf
            Case "3" '2016 選第3層
                Dim vCJOB_NO As String = cjob_value
                sql &= " SELECT CJOB_UNKEY ,CJOB_TYPE ,concat(CJOB_NO,'.',JOB_NO,CJOB_NAME) jobName,CJOB_NO" & vbCrLf
                sql &= " FROM SHARE_CJOB" & vbCrLf
                sql &= $" WHERE CJOB_NO IS NOT NULL AND JOB_NO IS NOT NULL AND CYEARS='2019' AND CDEL IS NULL AND CJOB_NO='{vCJOB_NO}'"
                sql &= " ORDER BY CJOB_NO ,JOB_NO" & vbCrLf
        End Select
        If sql = "" Then Return obj
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        With obj
            .DataSource = dt
            .DataTextField = "jobName"
            .DataValueField = "CJOB_UNKEY"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            'If TypeOf obj Is DropDownList Then .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        Return obj
    End Function

    '依 CJOB_UNKEY 回傳 CJOB_TYPE
    Function GetCJobTypeVal(ByVal cjobUnkey As String) As DataRow
        Dim rst As DataRow = Nothing
        cjobUnkey = TIMS.ClearSQM(cjobUnkey)
        If cjobUnkey = "" Then Return rst

        Select Case str_SHARECJOB_YEAR
            Case TIMS.cst_SHARE_CJOB_2016 '啟動2019-命名2016(cjobUnkey 舊轉新)
                Dim sqlx As String = $" SELECT UNKEY2 FROM SHARE_CJOB_REL WHERE UNKEY1='{cjobUnkey}'"
                Dim dr1 As DataRow = DbAccess.GetOneRow(sqlx, objconn)
                If Not dr1 Is Nothing Then cjobUnkey = Convert.ToString(dr1("UNKEY2"))
            Case Else '(不動)
        End Select

        Dim sql As String = ""
        sql &= " SELECT cjob_type ,cjob_no ,job_no ,cyears ,CJOB_UNKEY"
        sql &= $" FROM SHARE_CJOB WHERE CJOB_UNKEY='{cjobUnkey}'"
        rst = DbAccess.GetOneRow(sql, objconn)
        Return rst
    End Function

#End Region

#Region "2014" '第1代通俗職類

    '下拉選擇後動作。
    Private Sub CJOB_TYPE_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CJOB_TYPE.SelectedIndexChanged
        '依 CJOB_TYPE 設計 DropDownList  
        CJOB_NO = GetCJobUNKEYNO(CJOB_NO, CJOB_TYPE.SelectedValue, "")
    End Sub

    'Private Sub Button1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.ServerClick
    '    Common.RespWrite(Me, "<script>window.close();</script>")
    'End Sub

#End Region

    Protected Sub DdlCJOB16A1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCJOB16A1.SelectedIndexChanged
        ddlCJOB16A2.Items.Clear()
        ddlCJOB16A3.Items.Clear()
        ddlCJOB16A2 = GetCJobUNKEYNO(ddlCJOB16A2, ddlCJOB16A1.SelectedValue, "2")
    End Sub

    Protected Sub DdlCJOB16A2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCJOB16A2.SelectedIndexChanged
        'ddlCJOB16A2.Items.Clear()
        ddlCJOB16A3.Items.Clear()
        ddlCJOB16A3 = GetCJobUNKEYNO(ddlCJOB16A3, ddlCJOB16A2.SelectedValue, "3")
    End Sub
End Class
