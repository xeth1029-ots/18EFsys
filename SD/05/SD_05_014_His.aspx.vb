Partial Class SD_05_014_His
    Inherits AuthBasePage

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
        '檢查Session是否存在 End

        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow

        Dim rqtSOCID As String = TIMS.ClearSQM(Request("SOCID"))
        sql = "SELECT NAME,STUDENTID FROM dbo.VIEW_STUDENTBASICDATA WHERE SOCID='" & rqtSOCID & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then Exit Sub
        Name.Text = dr("Name")
        StudentID.Text = Right(dr("StudentID"), 2)

        sql = "SELECT * FROM dbo.STUD_TURNOUT2 WHERE SOCID='" & rqtSOCID & "'"
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGrid1.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count = 0 Then Exit Sub

        DataGrid1.Visible = True
        msg.Text = ""
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

End Class
