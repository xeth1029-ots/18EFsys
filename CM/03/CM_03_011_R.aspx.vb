Partial Class CM_03_011_R
    Inherits AuthBasePage

    'Const cst_sql_1 As String = "sql_1" '只要組合sql 
    'Const cst_sql_2 As String = "sql_2" '組合sql，要產生查詢
    Const cst_vsSqlString As String = "SqlString" 'CM_03_011 共用
    'Session(cst_vsSqlString) ViewState(cst_vsSqlString) ViewState("SqlString")
    'Dim connString = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在--------------------------End

        Dim sql As String = ""
        sql = Convert.ToString(Session(cst_vsSqlString))
        Dim sCmd As New SqlCommand(sql, objconn)

        Try
            'Dim conn As New SqlConnection
            'TIMS.TestDbConn(Me, conn, True)
            Label1.Text = (sm.UserInfo.Years - 1911)
            Session(cst_vsSqlString) = Nothing
            PrintDate.Text = Now.Date
#Region "(No Use)"

            'If Request("SearchPlan") = "W" Then
            '    lb_Plan.Text = "產業人才投資方案(提升勞工自主學習計畫)"
            'Else
            '    lb_Plan.Text = "產業人才投資方案(產業人才投資計畫)"
            'End If
            'With da
            '    .SelectCommand = New SqlCommand(sql, objconn)
            'End With
            'da.Fill(dt)

#End Region
            Dim dt As New DataTable
            'Dim da As New SqlDataAdapter
            With sCmd
                .Parameters.Clear()
                dt.Load(.ExecuteReader())
            End With
            'If conn.State = ConnectionState.Open Then conn.Close()
            'SQL語法查詢 (table@DataTable1)
            CM_03_011.CreateData(dt, Request("X"), Request("Y"), Request("YText"), DataTable1, objconn)
            'conn.Close()
            'da.Dispose()
        Catch ex As Exception
            Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg += "/*  sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            'Common.RespWrite(Me, "發生錯誤：" & ex.Message.ToString)
            ' Me.RegisterStartupScript("errMsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
        End Try
    End Sub
End Class
