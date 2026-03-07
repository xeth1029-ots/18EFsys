Partial Class CheckID
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        Call ChkACCOUNT10()
    End Sub

    Sub ChkACCOUNT10()
        Dim strScript As String = ""
        Dim RqID As String = Convert.ToString(Me.Request("id"))
        Dim vsRequestID As String = TIMS.ClearSQM(RqID)
        If RqID = "" Then
            lblMsg.Text = "請輸入帳號!!!"
            strScript = "<script>alert('請輸入帳號!!!');window.close();</script>"
            Page.RegisterStartupScript("", strScript)
            Exit Sub
        End If
        If RqID <> "" AndAlso vsRequestID <> RqID Then
            lblMsg.Text = "請重新輸入帳號!!!"
            strScript = "<script>alert('請重新輸入帳號!!!');window.close();</script>"
            Page.RegisterStartupScript("", strScript)
            Exit Sub
        End If
        vsRequestID = UCase(vsRequestID) '轉大寫
        If RqID = "" OrElse vsRequestID = "" Then
            lblMsg.Text = "請輸入帳號!!!"
            strScript = "<script>alert('請輸入帳號!!!');window.close();</script>"
            Page.RegisterStartupScript("", strScript)
            Exit Sub
        End If

        Dim sqlstr As String = ""
        sqlstr = "SELECT UPPER(ACCOUNT) ACCOUNT1 FROM AUTH_ACCOUNT WHERE UPPER(ACCOUNT) =@ACCOUNT"
        Dim sCmd As New SqlCommand(sqlstr, objconn)
        Dim dt1 As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value = vsRequestID '大寫判斷
            dt1.Load(.ExecuteReader())
        End With

        Dim flagname As Boolean = False '沒人使用
        If dt1.Rows.Count > 0 Then
            For Each dr1 As DataRow In dt1.Rows
                '轉換大寫後判斷
                If Convert.ToString(dr1("ACCOUNT1")) = vsRequestID Then
                    flagname = True '有人使用
                    'strScript += "opener.document.form1.nameid.value='';opener.document.form1.nameid.focus();"
                    lblMsg.Text = "該帳號已有人使用!!!"
                    strScript = "<script>alert('該帳號已有人使用!!!');window.close();</script>"
                    Page.RegisterStartupScript("", strScript)
                    Exit For
                End If
            Next
        End If

        If Not flagname Then
            'opener.document.form1.userpass.focus();
            lblMsg.Text = "您可以使用該帳號!!!"
            strScript = "<script>alert('您可以使用該帳號!!!');window.close();</script>"
            Page.RegisterStartupScript("", strScript)
            Exit Sub
        End If
    End Sub
End Class