Partial Class SD_15_024_R
    Inherits AuthBasePage

    'Const cst_sql_1 As String = "sql_1" '只要組合sql 
    'Const cst_sql_2 As String = "sql_2" '組合sql，要產生查詢
    'Session(cst_vsSqlString) ViewState(cst_vsSqlString) ViewState("SqlString")
    'Dim connString = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
    'SD_15_024 / CM_03_011 共用
    Const cst_vsSqlString As String = "SqlString"
    Const cst_vs_parms1 As String = "vs_parms1"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        Dim sql As String = Convert.ToString(Session(cst_vsSqlString))
        Dim parms1 As Hashtable = Session(cst_vs_parms1)
        Dim rq_X As String = TIMS.ClearSQM(Request("X"))
        Dim rq_Y As String = TIMS.ClearSQM(Request("Y"))
        Dim rq_YText As String = TIMS.ClearSQM(Request("YText"))

        Try
            'Dim conn As New SqlConnection
            'TIMS.TestDbConn(Me, conn, True)
            Label1.Text = (sm.UserInfo.Years - 1911)
            Session(cst_vsSqlString) = Nothing
            PrintDate.Text = Now.Date

            Dim dt As DataTable = Nothing
            dt = DbAccess.GetDataTable(sql, objconn, parms1)

            'If conn.State = ConnectionState.Open Then conn.Close()
            'SQL語法查詢 (table@DataTable1)
            SD_15_024.CreateData(dt, rq_X, rq_Y, rq_YText, DataTable1, objconn)
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
