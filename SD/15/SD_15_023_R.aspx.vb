Partial Class SD_15_023_R
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁 '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        Dim ReqSearchPlan As String = TIMS.ClearSQM(Request("SearchPlan"))
        Dim ReqX As String = TIMS.ClearSQM(Request("X"))
        Dim ReqY As String = TIMS.ClearSQM(Request("Y"))
        Dim ReqYText As String = TIMS.ClearSQM(Request("YText"))

        Try
            Label1.Text = (sm.UserInfo.Years - 1911) '民國年度

            '54:充電起飛計畫（在職）判斷方式 '28:產業人才投資方案
            Dim sTitle As String = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
            If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                Select Case Convert.ToString(ReqSearchPlan)
                    Case "G", "W"
                        sTitle &= "（" & TIMS.Get_PName28(Me, ReqSearchPlan, objconn) & "）"
                End Select
            End If
            lb_Plan.Text = sTitle
            PrintDate.Text = Now.Date

            If Session("SD_15_023_SqlStr") Is Nothing Then
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
            End If
            If Session("SD_15_023_parms") Is Nothing Then
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
            End If

            Dim dt As DataTable = Nothing
            Dim parms As Hashtable = New Hashtable()
            Dim sql As String = Session("SD_15_023_SqlStr") '= ViewState("SD_15_023_SqlStr") '= Sql
            parms = Session("SD_15_023_parms") '= ViewState("SD_15_023_parms") '= parms
            'sql = Session("SqlString") 'Session("SqlString") = Nothing
            dt = DbAccess.GetDataTable(sql, objconn, parms)
            If dt Is Nothing Then
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
            End If
            If dt.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
            End If
            Call SD_15_023.CreateData(dt, ReqX, ReqY, ReqYText, DataTable1, objconn)
            'conn.Close() 'da.Dispose()
        Catch ex As Exception
            'strErrmsg &= "/*  ex.ToString: */" & vbCrLf
            'strErrmsg &= ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Dim strErrmsg As String = TIMS.GetErrorMsg(Page) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)
            ' Me.RegisterStartupScript("errMsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
        End Try
    End Sub

End Class
