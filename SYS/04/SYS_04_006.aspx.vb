Partial Class SYS_04_006
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線

        If Not IsPostBack Then
            create()
        End If
        Button1.Attributes("onclick") = "return check_data();"
    End Sub

    Sub create()
        Dim sql As String = "SELECT * FROM SYS_DAYS"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then
            Days1.Text = dr("Days1").ToString
            Days2.Text = dr("Days2").ToString
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        'Dim conn As SqlConnection = DbAccess.GetConnection
        Dim sql As String = "SELECT * FROM SYS_DAYS"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
        Else
            dr = dt.Rows(0)
        End If
        dr("Days1") = Days1.Text
        dr("Days2") = Days2.Text
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)
        Common.MessageBox(Me, "儲存成功!")
    End Sub

End Class
