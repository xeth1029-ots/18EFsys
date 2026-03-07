Partial Class TC_01_005_del1
    Inherits AuthBasePage

    'Dim Re_courid As String
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

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim Re_courid As String = Request("courid")
        Re_courid = TIMS.ClearSQM(Re_courid)
        If Re_courid = "" Then Exit Sub

        Dim sqlstr As String = "DELETE COURSE_COURSEINFO WHERE COURID=@courid" ' & Re_courid
        Dim dCmd As New SqlCommand(sqlstr, objconn)
        Call TIMS.OpenDbConn(objconn)
        With dCmd
            .Parameters.Clear()
            .Parameters.Add("courid", SqlDbType.VarChar).Value = Re_courid
            .ExecuteNonQuery()
        End With
        'Response.Redirect("TC_01_005.aspx?ProcessType=del&ID=" & Request("ID") & "")
        Dim url1 As String = "TC_01_005.aspx?ProcessType=del&ID=" & Request("ID") & ""
        Call TIMS.Utl_Redirect(Me, objconn, url1)

        'Re_courid = Request("courid")
        'Dim sqldr As DataRow
        'Dim sqlAdapter As SqlDataAdapter
        'Dim sqlTable As New DataTable
        'Dim sqlstr_del = "delete Course_CourseInfo where courid=" & Re_courid
        'DbAccess.ExecuteNonQuery(sqlstr_del, objconn)
        'objconn.Close()
        'Response.Redirect("TC_01_005.aspx?ProcessType=del&ID=" & Request("ID") & "")
    End Sub

End Class
